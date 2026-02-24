# dm2xcod

[![PyPI](https://img.shields.io/pypi/v/dm2xcod.svg)](https://pypi.org/project/dm2xcod/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

DOCX to Markdown converter in Rust with Python bindings.

## Table of Contents

- [Why dm2xcod](#why-dm2xcod)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
- [CLI Reference](#cli-reference)
- [Architecture Overview](#architecture-overview)
- [Development](#development)
- [License](#license)

## Why dm2xcod

- Rust-based converter focused on predictable performance.
- Covers common DOCX structures: headings, lists, tables, notes, links, images.
- Supports image handling strategies: inline base64, save to directory, or skip.
- Exposes both CLI and Python (`PyO3`) entry points.
- Includes strict reference validation for footnote/comment/endnote integrity.

## Requirements

- Rust `1.75+` (building from source)
- Python `3.12+` (ABI3 wheel compatibility)

## Installation

### Python package

```bash
pip install dm2xcod
```

### CLI (cargo)

```bash
cargo install dm2xcod
```

### Rust library

```toml
[dependencies]
dm2xcod = "0.3"
```

## Quick Start

### CLI

```bash
# write to file
dm2xcod input.docx output.md

# print markdown to stdout
dm2xcod input.docx
```

### Python

```python
import dm2xcod

# path input
markdown = dm2xcod.convert_docx("document.docx")
print(markdown)

# bytes input
with open("document.docx", "rb") as f:
    markdown = dm2xcod.convert_docx(f.read())
```

### Rust

```rust
use dm2xcod::{ConvertOptions, DocxToMarkdown};

fn main() -> anyhow::Result<()> {
    let converter = DocxToMarkdown::new(ConvertOptions::default());
    let markdown = converter.convert("document.docx")?;
    println!("{}", markdown);
    Ok(())
}
```

## API Reference

### `ConvertOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `image_handling` | `ImageHandling` | `Inline` | Image output strategy |
| `preserve_whitespace` | `bool` | `false` | Preserve original spacing more strictly |
| `html_underline` | `bool` | `true` | Use HTML tags for underline output |
| `html_strikethrough` | `bool` | `false` | Use HTML tags for strikethrough output |
| `strict_reference_validation` | `bool` | `false` | Fail on unresolved note/comment references |

`ImageHandling` variants:

- `ImageHandling::Inline`
- `ImageHandling::SaveToDir(PathBuf)`
- `ImageHandling::Skip`

Example with non-default options:

```rust
use dm2xcod::{ConvertOptions, DocxToMarkdown, ImageHandling};

fn main() -> Result<(), dm2xcod::Error> {
    let options = ConvertOptions {
        image_handling: ImageHandling::SaveToDir("./images".into()),
        preserve_whitespace: true,
        html_underline: true,
        html_strikethrough: true,
        strict_reference_validation: true,
    };

    let converter = DocxToMarkdown::new(options);
    let markdown = converter.convert("document.docx")?;
    println!("{}", markdown);
    Ok(())
}
```

### Advanced: Custom extractor/renderer injection

`DocxToMarkdown::with_components(options, extractor, renderer)` lets you replace the default pipeline.

```rust
use dm2xcod::adapters::docx::AstExtractor;
use dm2xcod::converter::ConversionContext;
use dm2xcod::core::ast::{BlockNode, DocumentAst};
use dm2xcod::render::Renderer;
use dm2xcod::{ConvertOptions, DocxToMarkdown, Result};
use rs_docx::document::BodyContent;

#[derive(Debug, Default, Clone, Copy)]
struct MyExtractor;

impl AstExtractor for MyExtractor {
    fn extract<'a>(
        &self,
        _body: &[BodyContent<'a>],
        _context: &mut ConversionContext<'a>,
    ) -> Result<DocumentAst> {
        Ok(DocumentAst {
            blocks: vec![BlockNode::Paragraph("custom pipeline".to_string())],
            references: Default::default(),
        })
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct MyRenderer;

impl Renderer for MyRenderer {
    fn render(&self, document: &DocumentAst) -> Result<String> {
        Ok(format!("blocks={}", document.blocks.len()))
    }
}

fn main() -> Result<()> {
    let converter = DocxToMarkdown::with_components(
        ConvertOptions::default(),
        MyExtractor,
        MyRenderer,
    );
    let output = converter.convert("document.docx")?;
    println!("{}", output);
    Ok(())
}
```

### Python API

- `dm2xcod.convert_docx(input: str | bytes) -> str`
- Current Python entry point uses default conversion options.

## CLI Reference

```text
dm2xcod <INPUT> [OUTPUT] [--images-dir <DIR>] [--skip-images]
```

| Argument/Option | Description |
|---|---|
| `<INPUT>` | Input DOCX path (required) |
| `[OUTPUT]` | Output Markdown path (optional, otherwise stdout) |
| `--images-dir <DIR>` | Save extracted images to a directory |
| `--skip-images` | Skip image extraction/output |

## Architecture Overview

Conversion pipeline:

1. Parse DOCX (`rs_docx`)
2. Build conversion context (relationships, numbering, styles, references, image strategy)
3. Extract AST via adapter (`AstExtractor`)
4. Validate references (optional strict mode)
5. Render final markdown via renderer (`Renderer`)

Project layout:

```text
src/
  adapters/      # Input adapters (DOCX -> AST extraction boundary)
  core/          # Shared AST/model types
  converter/     # Orchestration and conversion context
  render/        # Markdown rendering + escaping
  lib.rs         # Public API (Rust + Python bindings)
  main.rs        # CLI entrypoint
```

## Development

### Build from source

```bash
# Rust library/CLI
cargo build --release

# Python extension in local env
pip install maturin
maturin develop --features python
```

### Test and lint

```bash
cargo test --all-features
cargo clippy --all-features --tests -- -D warnings
```

### Performance benchmark

```bash
# default: tests/aaa, 3 iterations, max 5 files
./scripts/run_perf_benchmark.sh

# custom: input_dir iterations max_files
./scripts/run_perf_benchmark.sh ./samples 5 10
```

Latest benchmark record (`2026-02-14`):

- Command: `./scripts/run_perf_benchmark.sh ./tests/aaa 10 10`
- Threshold gate: `./scripts/check_perf_threshold.sh ./output_tests/perf/latest.json 15.0` (`pass`)
- Environment: `macOS 26.2 (Darwin arm64)`, `rustc 1.92.0 (ded5c06cf 2025-12-08)`
- Result file: `output_tests/perf/latest.json`

```json
{"input_dir":"./tests/aaa","iterations":10,"files":2,"samples":20,"avg_ms":1.651,"min_ms":0.434,"max_ms":6.081,"total_ms":33.029,"overall_ms":33.034}
```

### Performance threshold gate

```bash
# fails if avg_ms exceeds threshold
./scripts/check_perf_threshold.sh ./output_tests/perf/latest.json 15.0
```

### Release notes

```bash
# auto-detect previous tag to HEAD
./scripts/generate_release_notes.sh

# explicit range and output file
./scripts/generate_release_notes.sh v0.3.9 v0.3.10 ./output_tests/release_notes.md
```

### API stability policy

See `docs/API_POLICY.md`.

## License

MIT
