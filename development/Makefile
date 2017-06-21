.PHONY: install

install: eol2eol.c
	mkdir -p ./bin/server
	cp -r ./server/ ./bin/server/server
	cp ./server.js ./bin/server/server.js
	cp escape_excel.pl ./bin/server
	cd ./bin/server
	npm install
	cd ../..