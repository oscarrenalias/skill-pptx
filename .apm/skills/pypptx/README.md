# pypptx — PowerPoint skill for agents

A Python CLI toolkit for reading, editing, and creating `.pptx` files,
designed for use by AI agents.

**Source repository:** https://github.com/oscarrenalias/skill-office
**Releases:** https://github.com/oscarrenalias/skill-office/releases
**Version:** 0.1.16

## What this skill provides

- Extract text from slides
- List, add, delete, and reorder slides
- Unpack / clean / repack `.pptx` files at the OPC/XML level
- Generate thumbnail grids for visual inspection
- Self-bootstrapping entry point (`pypptx.py`) with no external tooling required

See [SKILL.md](./SKILL.md) for full usage instructions.

## Installation

### Via APM

```bash
apm install oscarrenalias/skill-office
```

### Via zip (manual)

Download `skills-v0.1.16.zip` from the
[releases page](https://github.com/oscarrenalias/skill-office/releases) and
unzip it into your skills directory:

```bash
# Claude Code
unzip skills-v0.1.16.zip -d ~/.claude/skills/

# Codex / other agents
unzip skills-v0.1.16.zip -d ~/.agents/skills/
```

The zip contains all skills in this package (`pypptx`, `pyxlsx`). Each skill
self-bootstraps its own `.venv` on first run — no additional setup required.

---

> **Note:** This folder is automatically maintained by the release workflow
> in the source repository. Do not edit files here directly — any changes
> will be overwritten on the next release.
