# consulting-deck

The deck skill for the innovation bootcamp competition. Participants give it their
storyline, it gives them slides in the EAD template.

## Three modes

1. **Build the deck.** Participant pastes a numbered dot-dash storyline. Dot N becomes
   slide N. The skill checks it against the pitch structure, writes the slides and
   produces the deck.
2. **Challenge your thinking.** Jury-style questions on the problem, the three
   initiatives, the evidence and the ask. Asks rather than rewrites.
3. **Review.** Titles, spelling, numbers, sources, text that will not fit, leftovers.

## The pitch structure

Fixed, and enforced: cover, agenda or objective, situation, complication, problem
statement, three solutions, initiative detail (one slide each, a second when needed),
recap with impact, the ask. Backup slides sit after the ask.

The builder refuses to produce a deck that is missing a required slide, shows two
solutions instead of three, or leaves out what a jury will ask of an initiative.

## Across Claude, ChatGPT and Copilot

`SKILL.md` is self-contained, so the same file works in all three. It can be installed
as a plugin in Claude, or pasted or attached as a prompt in ChatGPT or Copilot.

- **Claude** can also run `scripts/build_deck.py` and hand back a finished `.pptx`.
- **ChatGPT** can run the same script if the participant uploads it with the template.
- **Copilot** composes the slides itself inside PowerPoint. The skill tells it which
  template layout to use and holds it to the word budgets, but it will not place shapes
  at exact coordinates the way the script does.

## Running the builder directly

```bash
python3 skills/consulting-deck/scripts/build_deck.py storyline.yaml \
  --profile skills/consulting-deck/profiles/ead.json \
  --out deck.pptx
```

`--check-only` validates and prints the titles without writing a file.
`--structure free` drops the pitch skeleton for a deck that is not a competition pitch.
`--template` points at a template elsewhere, for example a live copy in OneDrive.

Requires `python-pptx` and `pyyaml`.

## Images and charts

Images go on the three solutions boxes, on initiative detail slides, and on a visual
slide where a picture carries the point. `image:` takes a local path or a direct URL;
relative paths resolve from the storyline file. A missing or unreachable image leaves
a marked placeholder rather than failing the build.

The skill only takes web images from sources that license reuse, and puts the source
and licence in the slide's source line. No identifiable people, no third-party logos.

Charts are native PowerPoint charts, so they stay editable. Defaults follow consulting
practice: direct data labels, no value axis, no gridlines, legend only for two or more
series, four series maximum, never a second y-axis.

### A note on the chart colours

`chart_palette` in the profile is not the raw EAD accent set. Run as a categorical
series palette, the template accents fail: 004B87 sits outside the usable lightness
band, and 869CAE and 006269 read as grey. The brand green 007749 passes and is kept
as series one; the rest were re-stepped until all six colour checks pass, including
colour-blind separation. The fourth slot warns on contrast, which is why data labels
are on by default.

## Type sizes

All sizes live in `profiles/ead.json` under `type`, so the deck can be made larger or
smaller without touching code. Current settings are tuned for reading in a room: 22pt
titles, 15pt bullets, 12.5pt body in panels. The EAD master sets titles at 32pt for
topic labels, which no action title fits, so the profile overrides it.

## Known limits

- `.potx` files fail to open in python-pptx on content type. The builder patches this
  in a temporary copy and leaves the original untouched.
- Charts are native PowerPoint charts and stay editable, but only bar, column,
  horizontal bar, line and stacked column are supported.
- The two "Content Pattern" layouts are excluded: their title box is 7.16in wide
  because of the side graphic, which no action title fits.
