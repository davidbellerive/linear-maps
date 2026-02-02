# Illustrator Linear Transit Maps

This repository contains **programmatically generated linear transit maps** used by **Rail Fans Canada (RFC)** to illustrate rapid transit lines and station sequences across Canadian cities.

The maps are designed to be:
- visually consistent  
- SVG-based (infinite resolution)  
- CMS-friendly (Joomla-compatible)  
- usable on mobile without fighting site templates  

The core workflow uses **Adobe Illustrator scripting** to generate maps from a CSV source of truth, and publishes the resulting SVGs for direct embedding on the web.

---

## What this project does

- Generates **clean, minimal linear rail diagrams** (stations + line)
- Automatically handles:
  - long station names
  - angled or vertical labels
  - terminal padding
  - dynamic canvas height
- Exports **static SVG files** that:
  - require no external fonts
  - work without JavaScript
  - can be hosted on GitHub and embedded anywhere

These maps are intended for **editorial and reference use**, not geographic accuracy.

---

## Design principles

### 1. CSV is the source of truth

All maps are generated from a structured CSV containing:

- **system**  
  Used to generate the document name and folder hierarchy.

- **line ID**  
  Used to generate the document name and folder hierarchy.

- **line name**  
  Used to populate the map title.

- **years of operation**  
  Automatically inserted in parentheses. No validation is performed, allowing arbitrary text if desired.

- **RGB color**  
  Used for the line indicator.

- **station list** (pipe-separated)  
  No hard limit on character length. The map canvas is extended vertically as needed to prevent conflicts between the title, line, and station labels.  
  The canvas has been tested with station names up to **34 characters**. If additional space is required, the horizontal padding variable can be adjusted.

This structure makes it trivial to regenerate or update maps as networks evolve.

---

### 2. SVGs must be **self-contained**

Because the target site does not host custom fonts:

- all text is converted to **paths**
- no external font dependencies exist
- SVGs render identically across browsers and platforms

---

### 3. CMS-first constraints

The target environment (Joomla) imposes several practical limitations:

- external stylesheets may be blocked
- inline `<script>` tags are often stripped or escaped
- swipe gestures may be hijacked by templates (especially on mobile)

As a result:

- **no JavaScript is required** for normal usage
- all embeds rely on pure HTML with inline styles
- maps scale reasonably well on mobile and degrade gracefully

---

## Illustrator script

The generator script:

- creates one document per line
- dynamically sizes the artboard height based on label extents
- spaces stations evenly along a fixed-width baseline
- supports angled or vertical station labels
- exports SVGs using a consistent naming scheme

### Typical usage

1. Update `rfc-lines.csv`
2. Open Adobe Illustrator
3. Run `linear-map-generator.jsx`
4. Commit the exported SVGs

---

## Project status and disclaimer

This project was **not built to any formal software standard**, nor was it originally developed by someone with professional programming experience.

The code-base, including the Illustrator scripting structure, was developed with significant assistance from large language models (LLMs). Human oversight, testing, and iteration were applied, but the resulting code reflects a pragmatic, goal-driven approach rather than a formally engineered one.

- the codebase should not be considered production-grade software
- there are no guarantees regarding stability, performance, or long-term maintainability
- design decisions prioritize practical outcomes over architectural rigor

Over time, the intent is to **progressively refactor and replace LLM-generated code** with a manually written and better-structured codebase that is more suitable for use beyond the very specific one this initial version was meant to accomplish.

As such, this repository should be viewed as:
- a working tool
- an evolving experiment

Use, adapt, and contribute accordingly.
