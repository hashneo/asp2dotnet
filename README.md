# asp2dotnet

##To run##
```
node app.js --base <base ASP path> --page <start page> --out <target path> [--rabbit-hole] [--verbose] [--no-includes] [--overwrite]

--rabbit-hole = Follow lined ASP pages and try and process the whole site
--verbose = Print mor stuff to the screen
--no-includes = Ignore processing include files
--overwrite = By default the program won;t overwrite existinf aspx and vb files, this ignores that.

```
Make sure you have the latest version of node.js (0.12.0) else you will see

```
var sourcePath = path.parse( searchString );
                      ^
TypeError: Object #<Object> has no method 'parse'
```
