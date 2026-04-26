# Deployment Guide

## Before you deploy

Make sure your repo has:
- `streamlit_app.py`
- `requirements.txt`
- `.python-version`
- `secrets.toml.template`

Your entrypoint file is:

- `streamlit_app.py`

## Secrets

This app supports multiple LLM providers. You only need to provide the key(s) for the provider you want to use.

Example secrets:

```toml
ANTHROPIC_API_KEY = "your-anthropic-key-here"
OPENAI_API_KEY = "your-openai-key-here"
GOOGLE_API_KEY = "your-google-ai-studio-key-here"
GROQ_API_KEY = "your-groq-key-here"
