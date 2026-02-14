# Linear Maps — Line JSON Specification (v1)

This is the **new** line data format for the Illustrator `linear.jsx` script.

- **One file per line/variant**
- File extension: `.json`
- Colors are **HEX** only (`#RRGGBB`)
- No CSV compatibility or fallback

## Folder layout

By default, the script looks for a folder named `lines/` next to `linear.jsx`:

```
linear.jsx
linear-config.json
lines/
  line-1.json
  line-2.json
  ...
```

If `lines/` is not found, the script will prompt you to select a folder.

## Required fields (per file)

```json
{
  "system": "OC Transpo",
  "region": "Ottawa",
  "line_name": "Line 1 - Confederation Line",
  "years": "2019",
  "color": "#D30F1D",
  "stations": ["Tunney’s Pasture", "Bayview", "Pimisi"]
}
```

### Field meanings

- `system` *(string, required)*  
  Operator / network name. Used by `{system}` token.

- `region` *(string, required)*  
  City / metro / region. Mapped to internal `id` and used by `{id}` token.

- `line_name` *(string, required)*  
  Line/variant display name. Used by `{name}` token and default export filename.

- `years` *(string, optional)*  
  Display years (e.g., `2019`, `2001-2020`, `TBD`). Used by `{years}` and `{years_paren}`.

- `color` *(string, required)*  
  HEX stroke color in the form `#RRGGBB`.

- `stations` *(required)*  
  Either:
  - an array of station label strings (recommended), or
  - a single pipe-delimited string (e.g., `"A|B|C"`)

  Must contain **at least 2** stations.

## Notes

- Any “planned” station token extensions (`~`, `{}` icons, branching geometry, etc.) are **not** implemented in this v1 JSON switch.
