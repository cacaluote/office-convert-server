use bytes::Bytes;
use office_convert_client::{OfficeConvertClient, OfficeConvertLoadBalancer};
use std::sync::Arc;
use testcontainers::{
    core::{wait::HttpWaitStrategy, IntoContainerPort, WaitFor},
    runners::AsyncRunner,
    GenericImage, ImageExt,
};
use tokio::sync::Barrier;

#[tokio::test(flavor = "multi_thread")]
async fn attempt_spam() {
    let container = GenericImage::new("jacobtread/office-convert-server", "latest")
        .with_exposed_port(3000.tcp())
        .with_wait_for(WaitFor::http(
            HttpWaitStrategy::new("/status").with_expected_status_code(200u16),
        ))
        .with_env_var("RUST_LOG", "debug")
        .start()
        .await
        .unwrap();

    let host = container.get_host().await.unwrap();
    let host_port = container.get_host_port_ipv4(3000).await.unwrap();
    let client_url = format!("http://{host}:{host_port}");

    // Number of load balancers
    let sets = 1;

    // Number of "Convert this" per "set"
    let tasks = 10;

    // Simple thing for coordinating tasks "Keep the test running until this many things calls .wait()"
    let barrier = Arc::new(Barrier::new((tasks * sets) + 1));

    // Load the file to process
    let file = Bytes::from_static(include_bytes!("samples/sample.docx"));

    for set in 0..sets {
        // Setup a client to put into the load balancer
        let client = OfficeConvertClient::new(client_url.as_str()).unwrap();

        // Setup the load balancer
        let lb = OfficeConvertLoadBalancer::new([client]);
        let lb = Arc::new(lb);

        for task in 0..tasks {
            let lb = lb.clone();
            let barrier = barrier.clone();
            let file = file.clone();
            tokio::spawn(async move {
                println!("start job set = {set}, task = {task}");
                lb.convert(file).await.unwrap();
                println!("end job set = {set}, task = {task}");
                barrier.wait().await;
            });
        }
    }

    barrier.wait().await;
}
