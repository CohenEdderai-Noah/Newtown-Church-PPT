---
name: worship-song-ppt-editor
description: Rebuild church worship PowerPoint song sections in the style of the existing New City deck. Use when editing a worship-service `.pptx` to remove old songs, insert new songs from titles, images, or PDFs, preserve the current lyric-slide and `新城` separator style, split verse and chorus onto separate slides, and normalize punctuation out of displayed lyrics.
---

# Worship Song PPT Editor

Use this skill when the user wants the worship-song portion of a service deck replaced while keeping the deck's existing visual style.

## Workflow

1. Inspect the source deck before editing.
2. Reuse the existing lyric slide and `新城` separator slide as templates.
3. Replace the old song section instead of appending new songs after it.
4. Build the new song list as structured blocks, then run the bundled script to rebuild the section.
5. Verify the new slide order and lyric text in the output `.pptx`.

## Inputs

Support these user inputs:

- Song title only:
  Search the web for lyrics. Prefer official or primary sources when possible.
- Song image or PDF:
  Read the lyrics directly from the provided material. Use OCR if needed.
- Existing deck:
  Treat the current deck as the style template. Do not redesign the slides.

## Required Output Rules

- Keep the lyric slide style similar to the existing deck.
- Keep the `新城` separator slide from the deck and place it before the first song, between songs, and after the final song block.
- Remove the old song slides from the song section of the deck.
- Put verse and chorus on separate slides.
- Remove punctuation from displayed lyric lines and replace each punctuation mark with two spaces.
- Clear composer/footer text on lyric slides unless the user explicitly asks to keep it.

## Song Preparation

Prepare songs as a JSON file with this shape:

```json
{
  "songs": [
    {
      "title": "你真偉大",
      "blocks": [
        {
          "label": "verse 1",
          "lines": [
            "當我漫步 在林間樹下徘徊",
            "鳥語啾啾 柔美叫和樹梢"
          ]
        },
        {
          "label": "chorus",
          "lines": [
            "我心不禁歌頌我主我神",
            "你真偉大 你真偉大"
          ]
        }
      ]
    }
  ]
}
```

Guidelines:

- Keep one verse or one chorus per block.
- Split long songs into as many blocks as needed, but do not mix a verse and chorus in the same block.
- Preserve the user's intended ordering.
- Let the script normalize punctuation; do not pre-fill decorative punctuation.

## Run The Script

Use the bundled script:

```bash
python3 scripts/rebuild_song_section.py source.pptx songs.json output.pptx
```

What the script does:

- Detect the existing worship-song section in the New City service deck.
- Clone the deck's lyric-slide template and `新城` separator template.
- Remove the old song section from the presentation order.
- Insert the new song blocks with separators.
- Preserve the rest of the service deck.

## Verification

After running the script, verify:

- The first song section starts with `新城`.
- Each song block appears in the requested order.
- Verse and chorus are on separate slides.
- The old songs no longer appear in the presentation order.
- The next liturgy section, such as `主禱文`, still follows the song section correctly.

## Resources

- `scripts/rebuild_song_section.py`: Rebuild the song section by cloning the existing deck style and replacing old songs.
