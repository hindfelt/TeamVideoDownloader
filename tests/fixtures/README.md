The one-second MP4 fixtures contain generated black video and silent audio.
They contain no meeting data. Regenerate them from the repository root with:

```sh
ffmpeg -f lavfi -i color=c=black:s=64x64:r=10:d=1 -c:v libx264 -an -movflags empty_moov+frag_keyframe+default_base_moof -y tests/fixtures/video.mp4
ffmpeg -f lavfi -i anullsrc=r=48000:cl=mono -t 1 -c:a aac -vn -movflags empty_moov+frag_keyframe+default_base_moof -y tests/fixtures/audio.mp4
```

FFmpeg is only needed to regenerate these fixtures, not to run the tests.
