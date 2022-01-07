# Change Log
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [Unreleased]

## [0.6.0] - 2022-01-07
### Added
- Added a function to pick the correct platform user pattern
- Added a funciton to pick the correct platform icon 

- Added fields to the **rsmf_manifest.json**:
    + importance in the event
    + icon in the conversation

- Added CHANGELOG file to the project

### Changed
- Changed how event direction is evaluated. You no longer need to specify custodian

### Fixed
- Fixed conversation split by custom messages count value

### Revoved
- Removed the _CustodianID_ parameter
