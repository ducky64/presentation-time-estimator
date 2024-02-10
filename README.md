# presentation-time-estimator
Generates a timing estimate (WPM-based) for a presentation where the script is in the speaker notes

Works for PowerPoint slides (.pptx) only.
The speaker notes should only contain the narration script, this counts all the words in the speaker notes for timing.
By default, calibrated for 120wpm.

You'll need python-pptx:
```
pip install python-pptx
```

Usage:
```
python pptxExtractTiming.py [yourPresentation.pptx]
```

Example output:
```
**Introduction**: 0:00 + 229 words (1:54) + 0:02 breaks
**Background**: 1:56 + 185 words (1:32) + 0:01 breaks
**PlForPcbs**: 3:30 + 190 words (1:35) + 0:03 breaks
**Greyboxing**: 5:08 + 238 words (1:59) + 0:03 breaks
**Greyboxing II**: 7:10 + 166 words (1:23) + 0:03 breaks
**Discussion**: 8:36 + 174 words (1:27) + 0:00 breaks
**Future Work**: 10:03 + 179 words (1:29) + 0:00 breaks
**Conclusion**: 11:32 + 18 words (0:09) + 0:00 breaks

**TOTAL**: 11:41: 1379 words (11:29) + 0:12 breaks
```

Additionally generates a .md file with your script in one document.
