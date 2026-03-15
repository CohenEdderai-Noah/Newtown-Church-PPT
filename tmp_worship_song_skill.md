---
name: worship-song-ppt-editor
description: Rebuild church worship PowerPoint song sections in the style of the existing New City deck. Use when editing a worship-service `.pptx` to remove old songs, insert new songs from titles, images, or PDFs, preserve the current lyric-slide and `新城` separator style, ensure lyrics and `詞 / 曲 / 出處` metadata are in Traditional Chinese, split verse and chorus onto separate slides, and normalize punctuation out of displayed lyrics while keeping each slide concise.
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
  Search the web for lyrics and song metadata. Prefer official or primary sources when possible. Use Traditional Chinese lyrics; if the best source is Simplified Chinese, convert the final slide text to Traditional Chinese before building the deck.
- Song image or PDF:
  Read the lyrics and song metadata directly from the provided material. Use OCR if needed. Convert the final slide text to Traditional Chinese if the source is not already in Traditional Chinese.
- Existing deck:
  Treat the current deck as the style template. Do not redesign the slides.

## Required Output Rules

- Keep the lyric slide style similar to the existing deck.
- Keep the `新城` separator slide from the deck and place it before the first song, between songs, and after the final song block.
- Remove the old song slides from the song section of the deck.
- Always display lyrics in Traditional Chinese.
- Fill the text box between the title and lyrics with `詞：...  曲：...  出處：...`, and keep the labels and values in Traditional Chinese.
- Put verse and chorus on separate slides.
- Each lyric slide may contain at most one verse or one chorus. If a verse is too long for one slide, split it into multiple continuation blocks instead of adding a second verse.
- Each lyric slide must contain no more than five lyric lines.
- Each lyric line should generally contain eight or fewer Chinese characters. If the source line is longer, split it into shorter natural phrases while preserving meaning.
- Remove punctuation from displayed lyric lines and replace each punctuation mark with two spaces.
- Do not clear the intermediate metadata text box; populate it from the song data.

## Song Preparation

Prepare songs as a JSON file with this shape:

```json
{
  "songs": [
    {
      "title": "你真偉大",
      "lyricist": "佚名",
      "composer": "瑞典民謠",
      "source": "生命聖詩",
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

- Store `title`, `lyricist`, `composer`, `source`, and all lyric lines in Traditional Chinese.
- Keep one verse or one chorus per block.
- If a single verse or chorus exceeds the slide limits, split it into multiple blocks labeled as continuations, but do not mix different verses or a verse and chorus in the same block.
- Keep each block to five lyric lines or fewer.
- Keep each lyric line generally at eight or fewer Chinese characters by splitting long lines at natural phrase boundaries.
- Preserve the user's intended ordering.
- Provide `lyricist`, `composer`, and `source` for every song when available. If a value is unknown, keep the field present with an empty string.
- Ensure `lyricist`, `composer`, and `source` values are rendered in Traditional Chinese when a Chinese form is used.
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
- Write `詞：...  曲：...  出處：...` into the intermediate metadata placeholder on lyric slides.
- Preserve the rest of the service deck.

## Verification

After running the script, verify:

- The first song section starts with `新城`.
- Each song block appears in the requested order.
- Each lyric slide shows the correct `詞 / 曲 / 出處` line between the title and the lyrics, with Traditional Chinese labels and values where applicable.
- Verse and chorus are on separate slides.
- No lyric slide contains more than one verse or chorus block.
- No lyric slide contains more than five lyric lines.
- Lyric lines are generally eight or fewer Chinese characters after splitting.
- The old songs no longer appear in the presentation order.
- The next liturgy section, such as `主禱文`, still follows the song section correctly.

## Resources

- `scripts/rebuild_song_section.py`: Rebuild the song section by cloning the existing deck style and replacing old songs.
