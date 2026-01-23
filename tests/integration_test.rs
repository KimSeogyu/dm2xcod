use dm2xcod::{ConvertOptions, DocxToMarkdown};
use std::fs::{create_dir_all, read_dir, File};
use std::io::Write;
use std::process::Command;

#[test]
fn test_cli_help() {
    let output = Command::new(env!("CARGO_BIN_EXE_dm2xcod"))
        .arg("--help")
        .output()
        .expect("Failed to execute command");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Usage:"));
}

#[test]
fn verify_markdown_conversion() {
    let folders = ["./tests/aaa", "./tests/pandoc"];
    let output_dir = "./tests/output";
    create_dir_all(output_dir).expect("Failed to create output directory");

    let mut count = 0;
    let converter = DocxToMarkdown::new(ConvertOptions::default());

    for folder in folders {
        let dir = read_dir(folder).expect(&format!("Failed to read directory: {}", folder));
        for entry in dir {
            let entry = entry.unwrap();
            let path = entry.path();
            if path.is_file() && path.extension().map_or(false, |ext| ext == "docx") {
                println!("Verifying conversion: {:?}", path);
                let path_str = path.to_str().expect("Failed to convert path to string");
                match converter.convert(path_str) {
                    Ok(markdown) => {
                        let file_stem = path
                            .file_stem()
                            .expect("Failed to get file stem")
                            .to_str()
                            .expect("Failed to convert file stem to string");
                        let output_path = format!("{}/{}.md", output_dir, file_stem);
                        let mut file = File::create(&output_path)
                            .expect(&format!("Failed to create output file: {}", output_path));
                        file.write_all(markdown.as_bytes())
                            .expect("Failed to write to output file");
                        count += 1;
                    }
                    Err(e) => panic!("Failed to convert {:?}: {:?}", path, e),
                }
            }
        }
    }
    println!("Successfully verified {} files", count);
}
