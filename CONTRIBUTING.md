# Contributing

Thanks for your interest! This is a finished school/personal project kept as a portfolio piece, so changes are welcome but kept small: bug fixes, clarity, and compatibility with current Python versions.

## Prerequisites

- Python 3.8 or newer with Tkinter (bundled with the python.org installers; `sudo apt install python3-tk` on Debian/Ubuntu).
- The packages in `requirements.txt`.

## Check your change

```bash
pip install -r requirements.txt
python -m compileall -q .
python "Guess the number.py"   # play a quick game
```

The same checks run in CI on every pull request.

## Workflow

1. Create a branch from `main` named `type/short-topic` (for example `fix/flag-placement`).
2. Write commits following [Conventional Commits](https://www.conventionalcommits.org/) (`fix: …`, `feat: …`, `docs: …`).
3. Open a pull request using the template, and make sure CI is green.
4. Pull requests are merged with **squash and merge**.

## Versioning

The project follows [Semantic Versioning](https://semver.org/). User-visible changes go under `## [Unreleased]` in [CHANGELOG.md](CHANGELOG.md). A release is cut by tagging `main` with `vX.Y.Z`; the release workflow then publishes the GitHub Release from the changelog.

## Rules

- Never commit secrets, personal data or generated files.
- Be kind: see the [Code of Conduct](CODE_OF_CONDUCT.md).
