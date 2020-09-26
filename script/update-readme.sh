#!/bin/bash -eu

sed '/## Functions/q' README.md > README.tmp
npm run -s jsdoc2md up.js | tr -d '⇒' | sed '1d' >> README.tmp
mv README.tmp README.md
