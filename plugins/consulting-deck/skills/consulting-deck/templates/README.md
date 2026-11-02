# Template

`EAD Template.potx` is the template every deck is built into. It is the file EAD
provided. Do not edit it, do not change the master, and do not move the logo, the
green rule or the footer.

The builder uses it automatically. To point at a different copy, for example one EAD
has since updated:

    python3 scripts/build_deck.py storyline.yaml \
      --profile profiles/ead.json \
      --template "C:/path/to/EAD Template.potx" \
      --out deck.pptx
