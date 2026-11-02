# EAD innovation bootcamp: deck skill

A skill that helps you build the deck you will pitch to the jury. It works in Claude,
ChatGPT and Copilot.

It does three things:

1. **Build the deck.** You give it your storyline, it turns it into slides.
2. **Challenge your thinking.** It asks what the jury will ask.
3. **Review.** Typos, numbers, sources, and what to improve before you present.

## Install it

### Claude

In Claude, go to **Customize, then Plugins, then Add, then Add marketplace**, and paste:

```
simonedoro/EAD_Presentation-development
```

Then install **consulting-deck** from that marketplace and switch it on.

In Claude Code, the same thing from the terminal:

```bash
claude plugin install consulting-deck --marketplace simonedoro/EAD_Presentation-development
```

### Copilot

There is nothing to install. Copilot does not take plugins of this kind, so you give
it the instructions as a prompt.

1. Open the EAD template in PowerPoint first, so Copilot builds into it.
2. Open [`for-chatgpt-and-copilot/copilot-prompt.md`](for-chatgpt-and-copilot/copilot-prompt.md),
   copy everything below the line, and paste it into Copilot as your first message.

If pasting is awkward, use
[`copilot-prompt.docx`](for-chatgpt-and-copilot/copilot-prompt.docx) instead: save it
to your OneDrive and attach it in Copilot with `/`.

### ChatGPT

Attach
[`for-chatgpt-and-copilot/consulting-deck-prompt.md`](for-chatgpt-and-copilot/consulting-deck-prompt.md)
to a new chat, or paste its contents as your first message. It is the full version.
The shorter Copilot prompt works too.

## You also need the EAD template

The EAD PowerPoint template is not in this repository. Use the copy EAD gave you.

## Start it

Once it is installed, say:

> I want to build my bootcamp deck

and it will ask you for your storyline.

Bring your storyline as numbered dots. One dot is one slide:

```
1. [cover] Project title
2. [agenda] ...
3. [situation] Treated water goes to sea because the network cannot reach the districts that need it
   - treatment runs at 82% of capacity
   - reuse network reaches 11 of 34 districts
```

If you have not written one yet, say so and it will write it with you. Do not skip
that step. A deck built on a storyline you have not thought through falls apart on
the first jury question.

## What the deck looks like

The structure is fixed, and the skill will tell you if a piece is missing:

cover, agenda or objective, situation, complication, problem statement, three
solutions, one or two slides per initiative, a recap with impact, and the ask.
Anything else goes in speaker notes or in backup slides after the ask.
