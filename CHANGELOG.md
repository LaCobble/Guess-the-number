# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `requirements.txt` listing the scoreboard dependencies, with `xlrd` pinned below 2.0 (2.x cannot read `.xlsx`).
- Portfolio README, MIT license, contribution guide, code of conduct and security policy.
- Continuous integration (dependency install, byte-compilation, secret scan) and a release workflow.

### Fixed

- README: Tkinter is not installed with `pip`, and the script name is `Guess the number.py`.

## [1.0.0] - 2022-03-17

First tagged version: the game as built for the challenge.

### Added

- Tkinter game with Easy, Normal and Hard levels and "higher / lower" hints.
- Timed games and a per-level scoreboard stored in `classement.xlsx`.

[Unreleased]: https://github.com/LaCobble/Guess-the-number/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/LaCobble/Guess-the-number/releases/tag/v1.0.0
