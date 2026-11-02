# Put the EAD template here

The EAD PowerPoint template is not kept in this repository, because the repository is
public and the template is EAD's.

Save the file EAD gave you into this folder, keeping its name:

    EAD Template.potx

If you keep it somewhere else, point the builder at it:

    python3 scripts/build_deck.py storyline.yaml \
      --profile profiles/ead.json \
      --template "C:/path/to/EAD Template.potx" \
      --out deck.pptx
