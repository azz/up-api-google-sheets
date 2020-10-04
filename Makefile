install:
	npm install

test:
	npm test

build:
	./script/update-readme.sh

publish:
	npm run clasp push
