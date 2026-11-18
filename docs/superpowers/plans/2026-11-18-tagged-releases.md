# Tagged releases

Goal: derive one application and Excel version from `vA.B.C.D` / `VA.B.C.D`,
including the optional `-dev` suffix. Dev tags build the `dev` branch and publish
a prerelease; stable tags build the tagged commit. Record the actual source SHA.

Approach: prepare an ignored Python version module and Windows resource before
testing/building. Keep the source default for untagged builds. Both Excel writers
consume the same runtime version. The user selected EXE and ready XLSM in Release:
add a self-hosted Windows x64 runner with label `excel`. Pin it to the source SHA
resolved by the hosted EXE job. Publish only when both artifacts pass checks.

Constraints: retain the existing Tk capture workaround and database/data behavior.
Do not publish, create tags, or push during local implementation.

1. [ ] Add failing behavioral tests for strict tags, invalid Windows versions,
   generated CLI/resource versions, restamping, and XLSX metadata/footer.
2. [ ] Implement version preparation and connect EXE, GUI, XLSX and XLSM builders.
3. [ ] Connect tag selection, provenance, and GitHub Release publication after
   successful tests and build. Support uppercase V and dev prereleases.
4. [ ] Document use, run the full application suite, build a tagged EXE and XLSM
   locally, inspect versions in actual artifacts, and review the changes.

Review focus: invalid tags fail before publication; untagged builds reset stale
versions; dev provenance differs from event SHA; prereleases are not latest;
release jobs never run on pull requests; Excel versioning preserves data layout.
