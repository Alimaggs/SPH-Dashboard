# Roboto

`roboto-latin.woff2` is the Latin subset of Roboto, the typeface Material
Design 3 specifies. It is a variable font covering the weights the dashboard
uses (400 and 500) in a single 43 KB file.

It is embedded in the built page as a data URI rather than fetched from Google
Fonts, so the dashboard still loads in one request and the Content-Security-
Policy can keep denying every remote origin.

Licensed under the SIL Open Font License 1.1 — see `OFL.txt`, which is
redistributed with the font as the licence requires.
Copyright 2011 The Roboto Project Authors (https://github.com/googlefonts/roboto-classic)
