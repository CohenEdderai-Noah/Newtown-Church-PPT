# Song Spec

Use this file only when the JSON structure is unclear.

## Shape

```json
{
  "songs": [
    {
      "title": "輕輕聽",
      "blocks": [
        {
          "label": "verse 1",
          "lines": [
            "輕輕聽 我要輕輕聽",
            "我要側耳聽我主聲音"
          ]
        },
        {
          "label": "chorus",
          "lines": [
            "祢是大牧者 生命的主宰",
            "我一生只聽隨主聲音"
          ]
        }
      ]
    }
  ]
}
```

## Notes

- `label` is metadata for the human editor; the current script does not print it on the slide.
- Each `blocks[]` item becomes one lyric slide.
- The script removes punctuation from `lines[]` and replaces punctuation with two spaces.
