// This build script is used to embed resources like icons and manifests
// into the final executable on Windows.

fn main() {
    // We only need to run this script when targeting a Windows OS.
    if std::env::var("CARGO_CFG_TARGET_OS").unwrap() == "windows" {
        // Create a new WindowsResource object.
        let mut res = winres::WindowsResource::new();

        // 1. Set the application icon.
        // The path is relative to the project root (where Cargo.toml is).
        // Make sure you have an `assets/icon.ico` file.
        res.set_icon("assets/icon.ico");

        // 2. Set the application manifest.
        // This file tells Windows how to handle scaling, theming, and permissions.
        // Make sure you have an `assets/app.manifest` file.
        res.set_manifest_file("assets/app.manifest");

        // Compile the resources into the executable.
        // If this fails, the build process will stop with an error.
        if let Err(e) = res.compile() {
            eprintln!("Failed to compile windows resource: {}", e);
            std::process::exit(1);
        }
    }
}
