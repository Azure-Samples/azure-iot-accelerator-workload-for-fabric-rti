# Example device models

These model files are maintained as fixtures for this suite. Tests validate their
SHA-256 hashes before upload through the real browser file input.

| File | Scenario | SHA-256 |
| --- | --- | --- |
| `ThermoProX200-dtdl.json` | Nonlinked DTDL | `9b1d59c20f62401fc547eb3b292a28734021fe112733d50197b22524d922bd1b` |
| `ThermoProX200-wot-tm.json` | Combined and linked-only WoT | `169b0a51ea472bb3ad60256f587df0349bdb16bbbb55d52885064516c935a676` |

The DTDL bundle's root ID is `dtmi:example:ThermoPro_X200;1`.
The WoT bundle's root ID is `urn:example:ThermoProX200`, version `2.1.0`;
its inherited base and submodel documents are bundled in the same file.
No live device model reference is assigned just to satisfy a filter.

`src\models.ts` defines the explicitly selected fields and their expected roles
and native column types. The local simulator emits those fields, including DTDL
components and whole-value objects/arrays. The suite does not claim coverage of
every capability declared in the models.
