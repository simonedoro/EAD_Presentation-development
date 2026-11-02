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
2. Download [`copilot-prompt.docx`](for-chatgpt-and-copilot/copilot-prompt.docx),
   save it to your OneDrive, and attach it in Copilot with `/`.

Copilot does not read Markdown files. Use the Word version, or the
[PDF](for-chatgpt-and-copilot/copilot-prompt.pdf). If attaching does not work on your
machine, open
[`copilot-prompt.md`](for-chatgpt-and-copilot/copilot-prompt.md), copy everything
below the line, and paste it in as your first message. Pasting always works.

The Copilot version is deliberately short so it pastes reliably. Your situation and
complication slides are already filled in from the brief in your starter deck, so
Copilot does not need the case material to help you build.

### ChatGPT

Attach the file for your challenge from
[`for-chatgpt-and-copilot/cases/`](for-chatgpt-and-copilot/cases/) to a new chat, or
paste it in as your first message. Each one already contains your case material, so
you need nothing else.

[`consulting-deck-prompt.md`](for-chatgpt-and-copilot/consulting-deck-prompt.md) is
the same instructions without a case attached, if you want the generic version.

## Start from your group's starter deck

[`starter-decks/`](starter-decks/) has one deck per challenge, already in the EAD
template. Download the one for your challenge and open it in PowerPoint.

Each starter deck has the eleven slides of the pitch in order, with prompts where
your content goes. Your situation slide is already built as headline numbers, current
against target, and your complication slide as a chart, both from the challenge brief
with sources. Check them, change them if your work has moved on, and spend your time
on what you add to the problem rather than on restating it.

Image placeholders name what should be in the picture, for example "a two-compartment
kitchen bin in use". Find or take those three photos, or ask the skill to find
licensed ones for you.

`EAD executive starter deck.pptx` is the same pitch in the consulting exhibit layout:
evidence on the left, what it means on the right, the action title saying the point.
Use it if you want a deck that reads like a consulting deck. The placeholder numbers
are there to be replaced.

`EAD Template (empty).pptx` is the blank template, if you would rather start from
nothing.

## The template

The EAD template is included, both inside the plugin and as
`starter-decks/EAD Template (empty).pptx`.

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
