# pypptx — PowerPoint skill for agents

A Python CLI toolkit for reading, editing, and creating `.pptx` files,
designed for use by AI agents.

**Source repository:** https://github.com/oscarrenalias/skill-pptx
**Version:** 0.1.8

## What this skill provides

- Extract text from slides
- List, add, delete, and reorder slides
- Unpack / clean / repack `.pptx` files at the OPC/XML level
- Generate thumbnail grids for visual inspection
- Self-bootstrapping entry point (`pypptx.py`) with no external tooling required

See [SKILL.md](./SKILL.md) for full usage instructions.

## Installation

```bash
apm install oscarrenalias/skill-pptx
```

---

> **Note:** This folder is automatically maintained by the release workflow
> in the source repository. Do not edit files here directly — any changes
> will be overwritten on the next release.
