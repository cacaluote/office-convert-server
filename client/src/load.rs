use crate::{OfficeConvertClient, RequestError};
use bytes::Bytes;
use std::{sync::atomic::AtomicUsize, time::Duration};
use thiserror::Error;
use tokio::{
    sync::{Mutex, Notify},
    time::{sleep, timeout, Instant},
};
use tracing::{debug, error};

pub struct LoadBalancerTiming {
    /// Time in-between external busy checks
    pub retry_busy_check_after: Duration,
    /// Time to wait before repeated attempts
    pub retry_single_external: Duration,
    /// Timeout to wait on the notifier for
    pub notify_timeout: Duration,
}

impl Default for LoadBalancerTiming {
    fn default() -> Self {
        Self {
            retry_busy_check_after: Duration::from_secs(5),
            retry_single_external: Duration::from_secs(1),
            notify_timeout: Duration::from_secs(120),
        }
    }
}

#[derive(Debug, Error)]
pub enum LoadBalanceError {
    #[error("no servers available for load balancing")]
    NoServers,
}

struct LoadBalancedClient {
    /// The actual client
    client: OfficeConvertClient,

    /// Last time the server reported as busy externally
    busy_externally_at: Option<Instant>,
}

/// Round robbin load balancer, will pass convert jobs
/// around to the next available client, connections
/// will wait until there is an available client
pub struct OfficeConvertLoadBalancer {
    /// Available clients the load balancer can use
    clients: Vec<Mutex<LoadBalancedClient>>,

    /// Number of active in use clients
    active: AtomicUsize,

    /// Notifier for connections that are no longer busy
    free_notify: Notify,

    /// Timing for various actions
    timing: LoadBalancerTiming,

    /// Mutex used when checking for external blocking
    external_blocking_mutex: Mutex<()>,
}

impl OfficeConvertLoadBalancer {
    /// Creates a load balancer from the provided collection of clients
    ///
    /// ## Arguments
    /// * `clients` - The clients to load balance amongst
    pub fn new<I>(clients: I) -> Self
    where
        I: IntoIterator<Item = OfficeConvertClient>,
    {
        Self::new_with_timing(clients, Default::default())
    }

    /// Creates a load balancer from the provided collection of clients
    /// with timing configuration
    ///
    /// ## Arguments
    /// * `clients` - The clients to load balance amongst
    /// * `timing` - Timing configuration
    pub fn new_with_timing<I>(clients: I, timing: LoadBalancerTiming) -> Self
    where
        I: IntoIterator<Item = OfficeConvertClient>,
    {
        let clients = clients
            .into_iter()
            .map(|client| {
                Mutex::new(LoadBalancedClient {
                    client,
                    busy_externally_at: None,
                })
            })
            .collect::<Vec<_>>();

        Self {
            clients,
            free_notify: Notify::new(),
            active: AtomicUsize::new(0),
            timing,
            external_blocking_mutex: Default::default(),
        }
    }

    /// Checks if all client connections are blocked externally, used
    /// to handle the case when to not wait on notifiers
    pub async fn is_externally_blocked(&self) -> bool {
        // This guard is used to ensure we are the ONLY one checking for blocking clients
        // otherwise many threads will race the locking condition below starving the actual
        // converter from running
        let _guard = self.external_blocking_mutex.lock().await;

        self.clients
            .iter()
            // We are externally blocked if all clients are marked as
            // busy externally and none of the clients have locks held
            .all(|client| {
                client
                    .try_lock()
                    // We could obtain the lock (Its not in use) and the client was marked as busy externally
                    .is_ok_and(|client| client.busy_externally_at.is_some())
            })
    }

    pub async fn convert(&self, file: Bytes) -> Result<bytes::Bytes, RequestError> {
        let total_clients = self.clients.len();
        let multiple_clients = total_clients > 1;

        loop {
            for (index, client) in self.clients.iter().enumerate() {
                let mut client = match client.try_lock() {
                    Ok(value) => value,
                    // Server is already in use
                    Err(_) => continue,
                };

                let client = &mut *client;

                let now = Instant::now();

                if let Some(busy_externally_at) = client.busy_externally_at {
                    let since_check = now.duration_since(busy_externally_at);

                    // Don't check this server if the busy check timeout hasn't passed (only if we have multiple choices)
                    if since_check < self.timing.retry_busy_check_after && multiple_clients {
                        continue;
                    }
                }

                // Check if the server is busy externally (Busy outside of our control)
                let externally_busy = match client.client.is_busy().await {
                    Ok(value) => value,
                    Err(err) => {
                        error!("failed to perform server busy check at {index}: {err}");

                        // Mark erroneous servers as busy
                        true
                    }
                };

                // Store the busy state if busy
                if externally_busy {
                    debug!("server at {index} is busy externally");

                    client.busy_externally_at = Some(now);
                    continue;
                }

                // Clear external busy state
                client.busy_externally_at = None;

                debug!("obtained available server {index} for convert");

                // Increase active counter
                self.active
                    .fetch_add(1, std::sync::atomic::Ordering::SeqCst);

                let response = client.client.convert(file).await;

                // Notify waiters that this server is now free
                self.free_notify.notify_waiters();

                // Decrease active counter
                self.active
                    .fetch_sub(1, std::sync::atomic::Ordering::SeqCst);

                return response;
            }

            let active_counter = self.active.load(std::sync::atomic::Ordering::SeqCst);

            // Handle case where all clients are blocked externally, we won't be woken by any clients
            // in this case, so instead of waiting for the notifier we wait a short duration
            //
            // If number of active connections are zero we can assume we are blocked for some reason
            // likely an external factor, we would never get notified so we must poll instead?
            let externally_blocked = self.is_externally_blocked().await;
            if externally_blocked || active_counter < 1 {
                debug!("all servers are externally blocked, delaying next attempt");
                sleep(self.timing.retry_single_external).await;
                continue;
            }

            debug!("no available servers, waiting until one is available");

            // All servers are in use, wait for the free notifier, this has a timeout
            // incase a complication occurs
            _ = timeout(self.timing.notify_timeout, self.free_notify.notified()).await;
        }
    }
}
