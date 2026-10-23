---
tags: [skill-creator]
max_turns: 20
timeout_seconds: 600
allowed_tools: [Read, Glob, Grep, Skill, Write, Edit, Bash]
---

mise exec -- python app.py すると OPENAI_API_KEY が無いって言われる。mise はエラー出してないし、.mise.toml に _.file も書いてある。なんで？
