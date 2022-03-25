## Goals
This project was invented to help
extracting audio files from powerpoint files
(.pptx), to use seperately (for ex. in background with 2x) if needed

## How it works
### Requirement
- [x] It requires ffmpeg (or ffmpeg.exe, in windows) to be in your path
- [x] python3

### How to run
```bash
python extract_audio.py # windows
python3 extract_audio.py # *nix
```

It shows a file-dialog to select pptx files

It is required for your files to be in *pptx* format

Then it puts /MPEG-4 audio/ files in the same folder as
powerpoint files.

## TODOs
- [] choose output file name

- [] choose output path

- [] use tk for mode visual interaction
    - [] progress bar

- [] handle missing ffmpeg

- [] use ffmpeg binding for python

- [] write a blog post about this
