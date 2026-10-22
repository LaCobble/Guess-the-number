# Guess the Number

[![CI](https://github.com/LaCobble/Guess-the-number/actions/workflows/ci.yml/badge.svg)](https://github.com/LaCobble/Guess-the-number/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/python-3.8%2B-3776AB?logo=python&logoColor=white)
![GUI](https://img.shields.io/badge/GUI-Tkinter-informational)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A Tkinter "higher or lower" game with three difficulty levels and a scoreboard saved to an Excel file, built under a two-hour time limit.

<p align="center">
  <img src="https://github.com/LaCobble/Guess-the-number/assets/71151090/178403b6-793a-4b31-8592-ebe187dc04a2" width="40%" alt="Guess the Number main window">
</p>

## What it does

The game picks a random number and you have to find it, with "higher" or "lower" as the only hints. The faster you find it, the better your score.

| Level  | Number to find                         |
|--------|----------------------------------------|
| Easy   | an integer from 1 to 999               |
| Normal | an integer from 1 to 9,999             |
| Hard   | a decimal from 0.01 to 9,999.99        |

When you win, you enter a name and your time is stored in `classement.xlsx`, one sheet per level. The **Classement** (scoreboard) button shows the saved scores for each level.

## Getting started

### Prerequisites

- Python 3.8 or newer with Tkinter. Tkinter ships with the python.org installers for Windows and macOS; on Debian/Ubuntu install it with `sudo apt install python3-tk` (it is **not** installed with `pip`).
- The packages listed in [`requirements.txt`](requirements.txt): `xlrd` 1.x (version 2 dropped `.xlsx` support), `xlsxwriter` and `openpyxl`.

### Install and run

```bash
git clone https://github.com/LaCobble/Guess-the-number.git
cd Guess-the-number
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python "Guess the number.py"
```

The scoreboard file `classement.xlsx` is created in the current folder after the first win; it is ignored by Git.

## Project structure

```text
.
├── Guess the number.py   # the whole game: menu, game window, scoreboard read/write
├── requirements.txt      # Excel libraries used for the scoreboard
└── README.md
```

## What this project demonstrates

- Building a small desktop GUI with Tkinter (windows, buttons, entry fields, pop-ups).
- Persisting data to a spreadsheet with `xlsxwriter` and reading it back with `xlrd`.
- Timing a game with `time.perf_counter`.
- Delivering a complete, playable program under a strict two-hour constraint.

## Context

A personal challenge from 2021: build a complete game with a graphical interface in under two hours. The code is kept as it was written, as a portfolio piece; only the repository around it (documentation, dependency list, CI) has been tidied up.

## Contributing

Issues and small pull requests are welcome: see [CONTRIBUTING.md](CONTRIBUTING.md) and the [Code of Conduct](CODE_OF_CONDUCT.md).

## License

[MIT](LICENSE) © LaCobble
