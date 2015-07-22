# asp2dotnet

##To run##
```
node app.js --base <root ASP directory> --page <start page> --out <target path> [--rabbit-hole] [--verbose] [--no-includes] [--overwrite]

--rabbit-hole = Follow linked ASP pages and try and process the whole site.
--verbose = Print more stuff to the screen
--no-includes = Ignore processing include files
--overwrite = By default the program won't overwrite existing generated aspx and vb files, this ignores that.

And example would be:
node app.js --base c:\inetpub\wwwroot --page default.asp --out c:\inetpub\asp.net --rabbit-hole
```
Make sure you have the latest version of node.js (0.12.0) else you will see

```
var sourcePath = path.parse( searchString );
                      ^
TypeError: Object #<Object> has no method 'parse'
```
