use crate::{OfficeConvertClient, RequestError};
use bytes::Bytes;
use std::time::Duration;
use thiserror::Error;
use tokio::{
    sync::{Mutex, MutexGuard, Semaphore, SemaphorePermit},
    time::{sleep, Instant},
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

impl LoadBalancedClient {
    async fn check_busy(&mut self) -> Result<(), RequestError> {
        // Check if the server is busy externally (Busy outside of our control)
        let externally_busy = self.client.is_busy().await?;

        if externally_busy {
            // Store the busy state if busy
            self.busy_externally_at = Some(Instant::now());
        } else {
            // Clear external busy state
            self.busy_externally_at = None;
        }

        Ok(())
    }
}

/// Round robbin load balancer, will pass convert jobs
/// around to the next available client, connections
/// will wait until there is an available client
pub struct OfficeConvertLoadBalancer {
    /// Available clients the load balancer can use
    clients: Vec<Mutex<LoadBalancedClient>>,

    /// Permit for each client to track number of currently
    /// used client and waiting for free clients
    client_permit: Semaphore,

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

        let total_clients = clients.len();

        Self {
            clients,
            client_permit: Semaphore::new(total_clients),
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
        let (client, _client_permit) = self.acquire_client().await;
        client.client.convert(file).await
    }

    /// Acquire a client, will wait until a new client is available
    async fn acquire_client(&self) -> (MutexGuard<'_, LoadBalancedClient>, SemaphorePermit<'_>) {
        loop {
            if let Some(result) = self.try_acquire_client().await {
                return result;
            }

            // Get number of active client permits
            let active_counter = self.clients.len() - self.client_permit.available_permits();

            // Handle case where all clients are blocked externally, we won't be woken by any clients
            // in this case, so instead of waiting for the notifier we wait a short duration
            //
            // If number of active connections are zero we can assume we are blocked for some reason
            // likely an external factor, we would never get notified so we must poll instead?
            if active_counter < 1 || self.is_externally_blocked().await {
                debug!("all servers are externally blocked, delaying next attempt");
                sleep(self.timing.retry_single_external).await;
            }
        }
    }

    /// Attempt to acquire a client that is ready to be used
    /// and attempt a conversion
    ///
    /// Provides a [ActiveClientPermit] when this permit is dropped
    /// other clients will be notified that the resource is available
    /// again for use
    async fn try_acquire_client(
        &self,
    ) -> Option<(MutexGuard<'_, LoadBalancedClient>, SemaphorePermit<'_>)> {
        // Acquire a permit to obtain a client
        let client_permit = match self.client_permit.acquire().await {
            Ok(value) => value,
            Err(_) => return None,
        };

        let single_client = self.clients.len() > 1;
        let available_clients = self
            .clients
            .iter()
            // Include index for logging
            .enumerate()
            // Filter to only clients that aren't in use
            .filter_map(|(index, client)| match client.try_lock() {
                Ok(client_lock) => Some((index, client_lock)),
                // Server is already in uses
                Err(_) => None,
            });

        for (index, mut client_lock) in available_clients {
            let client = &mut *client_lock;

            // If we have more than one client and this client was already checked for being busy earlier
            // then this client will be skipped and won't be checked until a later point
            if let (false, Some(busy_externally_at)) = (single_client, client.busy_externally_at) {
                let now = Instant::now();
                let since_check = now.duration_since(busy_externally_at);
                if since_check < self.timing.retry_busy_check_after {
                    continue;
                }
            }

            // Update client busy state
            if let Err(err) = client.check_busy().await {
                error!("failed to perform server busy check at {index}, assuming busy: {err}");

                // Erroneous clients are considered busy
                client.busy_externally_at = Some(Instant::now());
                continue;
            }

            // Client is busy (Externally)
            if client.busy_externally_at.is_some() {
                debug!("server at {index} is busy externally");
                continue;
            }

            debug!("obtained available server {index} for convert");
            return Some((client_lock, client_permit));
        }

        None
    }
}
