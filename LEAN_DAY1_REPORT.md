# Lean London — Day 1 Report

Date: 2026-09-03
Branch: `london-lean-core-2026-09-03`

## Objective

Strip London down to one backend, remove Power Automate from the core design, preserve the existing working candidate, consolidate Microsoft/Dropbox/OpenAI logic, and establish automated tests.

## Completed

- Preserved `london-production-candidate-2026-09-02` unchanged for rollback/reference.
- Created isolated lean branch `london-lean-core-2026-09-03`.
- Added one Render runtime entrypoint: `lean-server.js`.
- Changed the lean branch start command to the single backend while retaining `npm run start:legacy` for rollback testing.
- Added direct Microsoft Graph client for London inbox reading, message retrieval and send-mail actions.
- Added one centralized OpenAI Responses API client with robust output parsing.
- Added Dropbox client restricted to `DROPBOX_ROOT_PATH`.
- Added restart-safe message deduplication state.
- Added central `LondonCore` orchestration.
- Removed Power Automate from the lean runtime path; the new server polls London's mailbox directly.
- Added protected health/manual-poll endpoints.
- Added automated regression tests and a GitHub Actions workflow definition.
- Updated `.env.example` with placeholders only; no real secret values were added.

## Automated verification completed locally

- Syntax validation: PASS
- Lean regression tests: 8/8 PASS
- Tested OpenAI response parsing when a reasoning item precedes the assistant message: PASS
- Tested Dropbox root confinement: PASS
- Tested principal-vs-external sender separation: PASS
- Tested duplicate-message suppression: PASS
- Tested self-message suppression: PASS
- Tested secret-free health/configuration reporting: PASS

## Not yet production-tested

The lean branch has not yet been cut over to the current Render production service. Live Microsoft/Dropbox/OpenAI calls must be verified against the lean branch before production merge. The previous candidate remains available and untouched.
