# dm2xcod

[![PyPI](https://img.shields.io/pypi/v/dm2xcod.svg)](https://pypi.org/project/dm2xcod/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A high-performance DOCX to Markdown converter written in Rust, with Python bindings.

## Features

- **Fast & Efficient**: Written in Rust for maximum performance.
- **Rich Formatting**: Preserves bold, italic, underline, strikethrough, and more.
  - Improves bold/italic rendering using HTML tags (`<strong>`, `<em>`) for better compatibility.
- **Structure Preservation**: Handles heading hierarchy, lists (ordered/unordered), and tables.
- **Image Support**: Extracts and embeds images.
- **Localization**: Specialized support for Korean document headings (`--lang ko`).
- **Cross-Platform**: Pre-built wheels for macOS (Intel/Apple Silicon), Windows, and Linux.
- **Simple API**: Native Python bindings provided via PyO3.

## Requirements

- **Rust**: 1.75+ (for building from source)
- **Python**: 3.12+ (Universal ABI3 support - works with 3.12, 3.13, 3.14+, etc.)

## Installation

### Python

Install via pip:

```bash
pip install dm2xcod
```

### CLI

Install via cargo:

```bash
cargo install dm2xcod
```

### Rust Library

Add to your `Cargo.toml`:

```toml
[dependencies]
dm2xcod = "0.3"
```

## Usage

### CLI

```bash
dm2xcod input.docx output.md

# With Korean localization support
dm2xcod input.docx output.md --lang ko
```

### Python

```python
import dm2xcod

# Basic conversion
markdown = dm2xcod.convert_docx("document.docx")
print(markdown)

# With options (if applicable in future versions)
# markdown = dm2xcod.convert_docx("document.docx", image_dir="images")
```

### Rust

```rust
use dm2xcod::{DocxToMarkdown, ConvertOptions};

fn main() -> anyhow::Result<()> {
    let converter = DocxToMarkdown::new(ConvertOptions::default());
    let markdown = converter.convert("document.docx")?;
    println!("{}", markdown);
    Ok(())
}
```

## Development

### Build from Source

```bash
# Build Rust library/CLI
cargo build --release

# Development with Python
pip install maturin
maturin develop --features python
```

## License

MIT
