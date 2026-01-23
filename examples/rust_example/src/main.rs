use dm2xcod::{ConvertOptions, DocxToMarkdown, ImageHandling};
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: {} <input.docx> [output.md]", args[0]);
        std::process::exit(1);
    }

    let input = &args[1];
    let output = args.get(2);

    let options = ConvertOptions {
        image_handling: ImageHandling::Inline,
        ..Default::default()
    };

    let converter = DocxToMarkdown::new(options);

    match converter.convert(input) {
        Ok(markdown) => {
            if let Some(out_path) = output {
                std::fs::write(out_path, &markdown).expect("Failed to write output file");
                println!("✅ Converted '{}' to '{}'", input, out_path);
            } else {
                println!("{}", markdown);
            }
        }
        Err(e) => {
            eprintln!("❌ Error: {}", e);
            std::process::exit(1);
        }
    }
}
