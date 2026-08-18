# Open XML validator

This deployment-time tool pins `DocumentFormat.OpenXml` 3.3.0 and emits JSON
with the package part and XPath for every schema error. Build it during skill
installation; conversion jobs use the resulting DLL without package restore or
network access.

```bash
bash tools/openxml-validator/build.sh
```

The build script is the only step allowed to restore packages. Runtime
validation never calls restore or build; a missing DLL is a fail-closed
deployment error.

Generated PPTX files require zero errors. Authored source decks may retain
extension warnings in evidence, but only a repair-free PowerPoint render can
accept those inputs.
