var fs = require('fs');
var path = require('path-extra');
var mkdirp = require('mkdirp');
var streamBuffers = require("stream-buffers");
var argv = require('optimist').argv;
var uuid = require('node-uuid');
var isNumeric = require("isnumeric");

var variablesParser = require('./variables-parser');
var propertiesParser = require('./properties-parser');
var functionsParser = require('./functions-parser');
var sanitizer = require('./code-sanitizer');

var functionMap = {};   // Map of global functions which we need to tie back to the defining class later
var functionMapCache = {};
var writtenClasses = {};
var basePath;
var sourceFiles = [];

var writtenFiles = [];

String.prototype.endsWith = function (suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
};

String.prototype.equalsIgnoreCase = function (what) {
    if (what === undefined)
        return false;
    return this.toLowerCase() === what.toLowerCase();
};


String.prototype.splice = function (pos,size) {
    if (size == 0)
        return this;
    var a = this.substr(0,pos);
    var b = this.substr(pos + size + 1);
    return a + b;
};

String.prototype.replaceAt = function(index, character) {
    return this.substr(0, index) + character + this.substr(index+character.length);
};

String.prototype.functionReplace = function( _from, _to, _left, _right ) {

    var re = new RegExp('[^\.]\\b' + _from + '\\b\\s*\\(.*', 'gi');

    var t = this;

    t = t.replace( re, function(match, p1){

        if ( _left === undefined ) _left = '(';
        if ( _right === undefined ) _right = ')';

        var pos;

        while( (pos = match.search(re)) != -1 ) {
            pos++;
            var a = 0;
            match = match.substr(0,pos) + _to + match.substr( pos + _from.length);
            pos += _to.length;
            var start = -1;
            var end = -1;
            for (var i = pos; i < match.length; i++) {
                if (match[i] == '(') {
                    a++;
                    if (start == -1) start = i;
                }

                if (match[i] == ')') {
                    if (--a == 0) {
                        end = i;
                        break;
                    }
                }
            }
            if (start != -1 && end != -1) {
                var t = match.replaceAt(start, _left).replaceAt(end, _right);
                match = t;
            }
        }
        return match;
    });

    return t;
};


function replaceInlineCode(code) {
    var regEx = /<%[\s\n\t]*=\s*([\s\S]+?)[\s\n\t]*%>/gi;
    var match;

    var copy =  code;
    while (( match = regEx.exec(code) ) != null) {
        var func = match[1];
        func = func.replace(/""/g, '"');
        copy = copy.replace(match[1], func);
    }

    code = copy.replace(regEx, '" & $1 & "');

    return code;
}

//***********************************************************************************************
//
//  Function: processFile
//
//
//***********************************************************************************************
function processFile( entry, rabbitHoleMode, writeMode ) {

    var sourceFile = entry.in;

    var vbSourceFile = false

    if ( sourceFile !== undefined )
        vbSourceFile = path.parse(sourceFile).ext.toLowerCase() === '.vb';

    // Prevent files being processed twice

    if ( writeMode ){

        if ( writtenClasses[entry.class] ){
            //console.log( 'I have already processed file => ' + sourceFile + ', skipping' );
            return;
        }

        writtenClasses[entry.class] = true;
    }

    var fileHeader;
    var codeBlocks = [];
    var includeFiles = [];

    var data;

    if ( entry.data !== undefined )
        data = entry.data;

    if ( sourceFile !== undefined )
        data = fs.readFileSync(sourceFile);

    if ( !writeMode ){
        if ( argv.verbose && sourceFile !== undefined  )
            console.log( 'pre-reading and merging file => ' + sourceFile );
    }

    var sourcePath = sourceFile !== undefined ? path.dirname( sourceFile ) : undefined;
    var targetPath = entry.vb != undefined ? path.dirname(  entry.vb ) : undefined;

    if ( entry.data === undefined ){
        data = data.toString('utf8')

        // Firstly replace all CRLF with LF (Windows)
        data = data.replace(/\r\n/g,'\n');

        // Replace any remaining CR with LF (Mac OS 9)
        data = data.replace(/\r/g,'\n');

        // Now clean the rest since we are In Unix Mode. Get rid of VBScript _ line separators, fix string concats, code markers
        data = data.replace(/_\s*\n(\t*|\s*)/gi,'').replace(/&"/gi,'& "').replace(/"&/gi,'" &').replace(/<%\s+=/gi,'<%=');
    }

    var match;


    if ( entry.data === undefined ) {

        // If there is no ASP VBScript tag, we move to module mode and don't create an ASP file
        /*
        if (!data.match(/<%@.*%>/g)) {
            entry.aspx = null;
            entry.vb = entry.vb.replace('.aspx.vb', '.vb');
            var className = entry.name.toLowerCase().replace(/_/g, ' ').replace(/(\b[a-z](?!\s))/g, function (x) {
                return x.toUpperCase();
            }).replace(/ /g, '');

            var prefix = 'cls';
            entry.class = prefix + className;
        }
        */

        if (targetPath !== undefined) {
            mkdirp.sync(targetPath, function (err) {
                console.log('could not create dir => ' + targetPath);
                throw err;
            });
        }
    }

    var regEx;

    // There won't be any includes in data mode since they should have already be processed
    if ( entry.data === undefined ) {
        // Strip out includes
        regEx = /<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->/gi;

        var htmlIncludes = [];

        data = data.replace(regEx, function (match, p1, p2) {

            var matchFile = p2.replace(/\\/g, '/');

            var includeFile = includeFile = path.normalize(path.join(sourcePath, matchFile)).trim();

            if (!fs.existsSync(includeFile)) {
                console.log('WARNING: include file => ' + includeFile + ' does not exists, skipping');
                return '';
            }

            var file = includeFile;

            var parts = path.parse(file);
            var filePath = parts.dir;
            var fileName = parts.name;

            if (parts.ext.toLowerCase() !== '.asp') {
                htmlIncludes.push({'file': matchFile, 'tag': match[0]});
                return '';
            }

            var subPath = path.relative(sourcePath, filePath).toLowerCase();

            var outPath = path.normalize(path.join(targetPath, subPath));

            var className = fileName.replace(/_/g, ' ').replace(/(\b[a-z](?!\s))/g, function (x) {
                return x.toUpperCase();
            }).replace(/ /g, '');

            var prefix = 'cls';
            fileName = fileName.toLowerCase();
            var aspxFile = null; //path.join(outPath, fileName + '.aspx');
            var vbFile = path.join(outPath, fileName + '.vb');

            includeFiles.push({'file': includeFile, 'class': prefix + className});

            if (argv.includes !== false)
                processFile({
                    'type' : 'include',
                    'name': fileName,
                    'class': prefix + className,
                    'relative': subPath,
                    'in': file,
                    'aspx': aspxFile,
                    'vb': vbFile
                }, rabbitHoleMode, writeMode);

            return '';
        });

        // Inject whatever html the file has included
        for ( var i = 0 ; i < htmlIncludes.length ; i++ ){
            var htmlInclude = htmlIncludes[i];
            var contents = '<% Response.WriteFile ("' + htmlInclude.file+ '")" %>';
            data = data.replace( htmlInclude.tag, contents );
        }
    }

    // Process the include files first as we need a function map later
    var aspx = null;
    var vb = null;

    if ( writeMode ){
        if ( entry.aspx != null ){
            if ( argv.overwrite == undefined && fs.existsSync(entry.aspx) ){
                console.log( 'WARNING: aspx file already exists => ' + entry.aspx + ', skipping' );
            }else{
                console.log( 'writing aspx source file => ' + entry.aspx );
                aspx = fs.createWriteStream(entry.aspx);
                writtenFiles.push(entry.aspx);
            }
        }

        if ( argv.overwrite == undefined && fs.existsSync(entry.vb) ){
            console.log( 'WARNING: vb file already exists => ' + entry.vb + ', skipping' );
        }else{
            if ( entry.vb != null ) {
                var skipFile = false;
                if (fs.existsSync(entry.vb)) {
                    var fd = fs.openSync(entry.vb, 'r');
                    var buffer = new Buffer(256);
                    fs.readSync(fd, buffer, 0, buffer.length, 0);
                    skipFile = (buffer.toString('utf8').indexOf('\'ignore') == 0);
                    fs.closeSync(fd);
                }

                if (skipFile) {
                    console.log('\'ignore declaration found in vb source file => ' + entry.vb);
                } else {
                    console.log('writing vb source file => ' + entry.vb);
                    vb = fs.createWriteStream(entry.vb);
                    writtenFiles.push(entry.vb);
                }
            }else{
                console.log('writing class => ' + entry.class + ' to a temp stream');
                vb = new streamBuffers.WritableStreamBuffer({
                    initialSize: (100 * 1024),      // start as 100 kilobytes.
                    incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
                });
            }
        }
    }

    var os = new streamBuffers.WritableStreamBuffer({
        initialSize: (100 * 1024),      // start as 100 kilobytes.
        incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
    });

    // Look for Linked ASP files as we will add them to the list we need to process
    if (!vbSourceFile){

        var regEx = /"(([\.-a-zA-Z0-9/\-_]+?\.asp)(?!\w+)(\?*.*)?)"/gi;
        data = data.replace( regEx, function(match, p1, p2, p3, offset, string){
            if ( rabbitHoleMode && writeMode ){
                /*
                var p2Parts = path.parse(p2);
                if ( p2Parts.name.toLowerCase() === entry.name.toLowerCase() ){
                    return match;
                }else{
                */
                    var nextFile = path.join( basePath, p2 );
                    var  addFile = true;
                    for ( var x = 0 ; x < sourceFiles.length ; x++ ){
                        if ( nextFile.toLowerCase() === sourceFiles[x].toLowerCase() ){
                            addFile = false;
                            break;
                        }
                    }
                    if ( addFile ){
                        if ( !fs.existsSync(nextFile)  ){
                            //console.log( 'WARNING: asp file => ' + nextFile + ' is referenced in => ' + sourceFile + ' but it doesn\'t exists' );
                            return match;
                        }else{
                            sourceFiles.push( nextFile );
                        }
                    }
                    return match.replace(p2, p2.replace('.asp','.aspx'));

                //}

            }else{
                return match;
            }
        });

    }

    // Remove classes
    regEx = /((?!'(?:\n)|(?:\n))(?:\s*'.*?(?:\n))*(?:public|private)?\s*\b(class\b\s*(\w+)\s*(?:'.*)?(?:\n))([\s\S]*?)(?:end\s+(?:class)))/gmi;

    var remainingData = data;

    var classBlocks = '';
    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[1];
        var className = match[3];
        var classData = match[4];

        if ( className === 'cAuthor')
            __debug = 1;

        if ( entry.name !== 'vb6'){
            classData = processFile( { 'type' : 'class', 'name' : fileName, 'class' : className,  'data' : '<%\n\n' + classData + '\n\n%>' }, rabbitHoleMode, writeMode );
        }else{
            classData = codeBlock;
        }

        if ( writeMode ){
            classBlocks += classData;
        }

        remainingData = remainingData.replace(codeBlock, "");
    }

    if (!vbSourceFile) {

        function processHtmlData(htmlCode){
            var outData = '';
            htmlCode = htmlCode.replace(/"/g, '""'); //.replace(/\t||\n/gi, '');
            var regEx = /([^\n]+)/gi;

            while (( match = regEx.exec(htmlCode) ) != null) {
                var line = sanitizer.clean( replaceInlineCode(match[1]) );

                line = line.replace(/<%@.*%>/g, '');

                if (line.trim().length > 0)
                    outData += '\n\t\tResponse.Write ("' + line + '")';
            }

            return outData + '\n';
        }

        // Replace RUNAT Server Tag with VBScript Marker
        regEx = /<SCRIPT\s+LANGUAGE\s*=\s*"VBScript"\s+RUNAT\s*=\s*"Server">([\s\n\t]*(?=[^=|@])([\s\S]+)[\s\n\t]*)<\/SCRIPT>/gi;

        remainingData = remainingData.replace(regEx, function (match, p1) {
            return '<%\n' + p1 + '%>\n';
        });


        // Convert all inline HTML code to Response.Writes
        var codeRegEx = /(?:<%)[\s\n\t]*(?=[^=|@])([\s\S]+?)[\s\n\t]*(?:%>\s*)/gi;

        var startPos = 0;
        var endPos = 0;

        data = remainingData;
        while (( match = codeRegEx.exec(data) ) != null) {

            var codeBlock  = match[1];

            endPos = match.index;
            var htmlBlock = data.substr(startPos, endPos - startPos);

            startPos = endPos + match[0].length;
            endPos = startPos;

            os.write( processHtmlData( htmlBlock ) );
            os.write( codeBlock );
        }

        if ( endPos < data.length ){
            os.write( processHtmlData( data.substr( endPos ) ) );
        }

        // Get the stream data and convert it to a string
        data = os.getContents();
        remainingData = data.toString('utf8');
    }

    remainingData = remainingData.replace( /('\s*\$Header[\s\S]+?All\sRights\sReserved.)/gi, function(m,p1){
        fileHeader = p1;
        return '';
    });

    // First remove all Functions
    var result = functionsParser.parse( remainingData, argv.verbose );
    var functionBlocks = result.blocks;
    functionMap[entry.class] = result.map;
    functionMap[entry.class]['_Type'] = entry.type;

    // Next we remove all Properties (GET/LET)
    var result = propertiesParser.parse( result.data, argv.verbose );
    functionMap[entry.class]['_Properties'] = result.properties;

    var globalBlocks = '';

    // Finally get all of the Variables
    result = variablesParser.parse( result.data, argv.verbose );

    functionMap[entry.class]['_Variables'] = result.vars;
    functionMap[entry.class]['_Constants'] = result.consts;

    if ( result.unmatched.length > 0 ){
        console.log("WARNING: We had unmatched data from variables => " + result.unmatched);
    }

    remainingData = result.data;
    globalBlocks += result.unmatched;

    var constDeclStr = '';

    // Only do substitutions when we are writing
    if ( writeMode && vb != null ) {

        // Remove all extra comments and multiple new lines
        remainingData = remainingData.replace(/('.*\n(\n){2})/g, '\n').replace(/^\s*\n{2}/gm, '');

        function formatCommentBlock(data) {
            var commentBlock = '';
            for (var i = 0; i < data.length; i++) {
                if (data[i].trim().length > 0)
                    commentBlock += '\t\' ' + data[i].trim() + '\n';
            }
            if (commentBlock.length > 0)
                commentBlock = '\n' + commentBlock;

            return commentBlock;
        }

        for (var i = 0; i < functionMap[entry.class]['_Constants'].length; i++) {
            var constDecl = functionMap[entry.class]['_Constants'][i];
            var type = constDecl.type;
            if ( type === undefined ) type = "Object"
            var Line = '\t' + constDecl.visibility + ' Const ' + constDecl.var + ' As ' + type;
            Line += new Array(Math.max(0, 50 - Line.length)).join(' ') + ' = ' + constDecl.value;
            if (constDecl.comment !== undefined) {
                Line = formatCommentBlock(constDecl.comment) + Line;
            }
            constDeclStr += Line + '\n';
        }

        var varDelcStr = '';

        for (var i = 0; i < functionMap[entry.class]['_Variables'].length; i++) {
            var varDecl = functionMap[entry.class]['_Variables'][i];
            for (var x = 0; x < varDecl.length; x++) {
                var type = varDecl[x].type;
                if ( type === undefined ) type = "Object"
                var Line = '\t' + varDecl[x].visibility + ' Property ' + varDecl[x].var  + ' As ' + type;
                if (varDecl[x].value !== undefined)
                    Line += new Array(Math.max(0, 50 - Line.length)).join(' ') + ' = ' + varDecl[x].value;
                if (varDecl[x].comment !== undefined) {
                    Line = formatCommentBlock(varDecl[x].comment) + Line;
                }
                varDelcStr += Line + '\n';
            }
        }

        var usedModules = [];

        //***********************************************************************************************
        //
        //  Function: processFunctionMap
        //
        //
        //***********************************************************************************************
        function processFunctionMap(block) {

            var thisFunction = block.function;
            var localVariables = block.vars;

            var code = block.code;

            if (code.trim().length == 0)
                return code;

            var includeClasses =  [ entry.class ];

            for (var cls in functionMap){
                if ( cls === entry.name || cls.indexOf('page') == 0 )
                    continue;
                includeClasses.push( cls );
            }

            for (var a = 0 ; a < includeClasses.length ; a++ ){

                var cls = includeClasses[a];

                var varName = "_" + cls; //.substring(3);

                var inUse = false;

                //************************************************************
                // function doSubstitution
                //************************************************************
                function doSubstitution(_type, _what, _with, onMatch) {

                    var regEx;

                    //regEx = new RegExp('^\\n?\\s*((?=(?:public|private)?\\s*(sub|function))|.*?)((?:\\w+\\.)?\\b' + _what.name + '\\b).*$', 'gmi' )

                    regEx = new RegExp('.*?\\b((?:\\w+\\.)?' + _what.name + ')\\b', 'gi');

                    code = code.replace( regEx, function(m, p1, p2, p3, offset, string){

                        if ( _what.name === 'ID' && code.indexOf('edit_book.asp') > 0 )
                            _debug = 1;

                        //if ( m.indexOf('AssociateAuthorWithBook') > 0 )
                        //debugger;
                        // I have no idea why the above regex is ignoring
                        /*
                        if ( p1.match(/\b(sub|function)\b/gi) != null && p2 === undefined ){
                            return m;
                        }

                        if (p2 !== undefined){
                            return m;
                        }

                        p1 = p3;
*/
                        if ( !onMatch(p1) ){
                            return m;
                        }

                        if ( m.indexOf('"')!= -1 ){
                            /*
                             var r = new RegExp( '(["\'])(?:\\b(' + _what.name + ')\\b.*?)?\\1','gi' );
                             if ( ( (m2 = m.match(r))) != null ) {
                             if ( m2[0].indexOf(_what.name) != -1)
                             return m;
                             }*/
                        }

                        // Check for being in comments or Strings
                        if ( m.indexOf('\'')!= -1 ){
                            var _where =  m.search(new RegExp( '\\b' + _what.name + '\\b','gi' ));
                            if ( _where != -1 ){
                                for ( var i = _where ; i > 0 ; i-- ){
                                    if ( m[i] == '"' ){
                                        break;
                                    }
                                    if ( m[i] == '\'' ){
                                        return m;
                                    }
                                }
                            }
                        }



                        _what.hits += 1;
                        inUse = true;

                        if ( _what.parameters !== undefined && _what.parameters.length > 0 ){
                            if (  m.match( new RegExp( p1 + '\\s*\\(') ) == null ){
                                m = m.replace( new RegExp('^(.*)\\b' + p1 + '\\b\\s*(?!\\s*=)(.+?)(?=(?:$))', 'm'), function(m2, p2, p3){
                                    return p2 + p1 + '( ' + p3 + ' )';
                                });
                            }
                        }
                        return m.replace( new RegExp('\\b' + p1 + '\\b', 'gi'), function(m2){
                            return  _with;
                        });
                    });
                }

                function isReservedName(n){
                    n = n.toLowerCase();

                    if ( n === 'new')
                        return true;

                    if ( n === '_type')
                        return true;

                    return false;
                }

                if ( functionMap[cls]['_Type'] === 'class'){
                    doSubstitution('var', { 'name' : 'new\\s+' + cls }, 'New ' + cls + '( Page )', function(match){
                        if ( match.indexOf('.') > 0 ){
                            var whichClass = match.split('.')[0];
                            if ( !whichClass.equalsIgnoreCase( item.name) )
                                return false;
                        }
                        return true;
                    });
                }

                for (var name in functionMap[cls]) {

                    if ( thisFunction.name === 'OTRunProcReturnCodeWithTimeout' && name === 'OTRunProcReturnCode')
                        __Debug = 1;

                    if ( isReservedName(name) || ( thisFunction !== undefined && name.toLowerCase() === thisFunction.name.toLowerCase() ) )
                        continue;

                    if (name[0] === '_') {

                        functionMap[cls][name].forEach( function( item ){

                            //if ( item.name.toLowerCase() === 'id' && cls.toLowerCase() === 'cauthor' )
                            //   debugger;

                            if ( ( thisFunction !== undefined && item.name.toLowerCase() === thisFunction.name.toLowerCase() ) )
                                return;

                            if ( localVariables !== undefined && localVariables.contains(item.var) )
                                return;

                            var _with = varName + '.' + item.var;
                            if (cls === entry.class)
                                _with = 'Me.' + item.var;
                            if ( name === '_Constants')
                                _with = cls + '.' + item.var;

                            if ( item.visibility !== undefined && item.visibility.toLowerCase() === 'private')
                                return;
                            doSubstitution('var', item, _with, function(match){
                                //if ( match.indexOf('ID') > 0 && code.indexOf('skyblue') > 0)
                                //    debugger;
                                if ( match.indexOf('.') > 0 ){
                                    var whichClass = match.split('.')[0];
                                    if ( !whichClass.equalsIgnoreCase( item.name) )
                                        return false;
                                }
                                return true;
                            });
                        });

                    } else {
                        var _with = varName + '.' + name;
                        if (cls === entry.class)
                            _with = 'Me.' + name;
                        if ( functionMap[cls][name].type === 'class')
                            _with = cls + '.' + name
                        if ( functionMap[cls][name].global )
                            _with = cls + '.' + name
                        if ( functionMap[cls][name].visibility.toLowerCase() === 'private')
                            continue;
                        doSubstitution('function', functionMap[cls][name], _with, function(match){
                            //if ( match.indexOf('ID') > 0 && code.indexOf('skyblue') > 0)
                            //    debugger;
                            if ( match.indexOf('.') > 0 ){
                                var whichClass = match.split('.')[0];
                                if ( !whichClass.equalsIgnoreCase( functionMap[cls][name].name) )
                                    return false;
                            }
                            return true;
                        });
                    }
                }

                if (inUse){
                    if ( cls !== entry.class ) {
                        var modFound = false;
                        for (var i = 0; !modFound && i < usedModules.length; i++) {
                            modFound = usedModules[i].class === cls;
                        }
                        if (!modFound) {
                            usedModules.push({'class': cls, 'var': varName});
                        }
                    }
                }
            }

            return code;
        }

        var propertiesStr = '';

        for ( var propertyName in  functionMap[entry.class]['_Properties'] ){
            var classProperty =  functionMap[entry.class]['_Properties'][propertyName];

            if (typeof classProperty === "function") {
                continue;
            }
            var theVar = classProperty._Variable;

            if (theVar != undefined ){
                var Line = '\t' + theVar.visibility + ' ' + theVar.name + " As Object"
                if (theVar.comment !== undefined) {
                    Line = formatCommentBlock(theVar.comment) + Line;
                }
                varDelcStr += Line + '\n';
            }

            var getStr = '';
            var setStr = '';

            if ( classProperty.get !== undefined ){
                var theGet = classProperty.get;
                getStr = (theGet.visibility.toLowerCase() !== 'public' ? theGet.visibility : '') + ' Get' +
                    processFunctionMap(theGet) +
                    'End Get\n'
            }

            if ( classProperty.let !== undefined ){
                var theSet = classProperty.let;
                setStr = (theSet.visibility.toLowerCase() !== 'public' ? theSet.visibility : '') + ' Set(' +  theSet.setParam + ' As Object )' +
                    processFunctionMap(theSet) +
                    'End Set\n'
            }

            var readWriteOnly = ''

            if ( classProperty.get === undefined ){
                readWriteOnly = 'WriteOnly'
            }

            if ( classProperty.let === undefined ){
                readWriteOnly = 'ReadOnly'
            }

            Line = 'Public ' + readWriteOnly + ' Property ' + propertyName + '\n' +
                getStr +
                setStr +
                'End Property\n\n';

            //Line = formatCommentBlock(theVar.comment) + Line;

            propertiesStr += Line + '\n'
        }

        var functionBlocksCode = '';

        for ( var i = 0 ; i < functionBlocks.length ; i++ ){
            var functionBlock = functionBlocks[i];
            var code = vbSourceFile ? functionBlock.code : processFunctionMap(functionBlock);
            functionBlocksCode += functionBlock.function.signature + '\n' +
                                  code.trim() + '\n' +
                                  'End ' +  functionBlock.function.type + ' \n\n';
        }


        if ( !vbSourceFile ) {
            remainingData = processFunctionMap({ 'function' : { name : '', 'parameters' : undefined } , 'code' : remainingData } );
        }

        var dimModules = ''
        var newModules = ''

        for (var i = 0; i < usedModules.length; i++) {
            var varName = '_' + usedModules[i].var;
            dimModules += '\tPrivate ' + varName + ' As ' + usedModules[i].class + ' = Nothing\n';
            dimModules += '\tPrivate ReadOnly Property ' + usedModules[i].var + ' As ' + usedModules[i].class + '\n';
            dimModules += '\t\tGet\n';
            dimModules += '\t\tIf ' + varName + ' Is Nothing Then \n';
            dimModules += '\t\t' + varName + ' = createInstance( GetType(' + usedModules[i].class + ') )\n';
            dimModules += '\t\tEnd If\n';
            dimModules += '\t\tReturn ' + varName + '\n';
            dimModules += '\t\tEnd Get\n';

            dimModules += '\tEnd Property\n';

            //newModules += '\t\t' + usedModules[i].var + ' = createInstance( GetType(' + usedModules[i].class + ') )\n';
        }

        var writeClass = true;


        if ( constDeclStr.trim() === '' &&
             varDelcStr.trim() === '' &&
             propertiesStr.trim() === '' &&
             dimModules.trim() === '' &&
             globalBlocks.trim() === '' &&
             functionBlocksCode.trim() === '' &&
             remainingData.trim() === '' )
            writeClass = false;

        if (aspx != null) {
            aspx.write( '<%@ Page Language="VB" AutoEventWireup="true" CodeBehind="' + entry.name + '.aspx.vb" Inherits="' + argv.project + '.' + entry.class + '" %>' )
        }

        if ( writeClass ) {


            if (aspx != null) {
                vb.write('Public Class ' + entry.class + '\n\n');
                vb.write('\tInherits AspPage' + '\n');
                vb.write('\n');
            } else {
                vb.write('Public Class ' + entry.class + '\n\n');
                vb.write('\tInherits PageClass' + '\n')

                vb.write('\n');
            }

            if (fileHeader !== undefined) {
                vb.write('\t#Region "Original Header"\n');
                vb.write(fileHeader + '\n');
                vb.write('\t#End Region\n\n');
            }

            if (constDeclStr.length > 0) {
                vb.write('\t#Region "Global Constants"\n');
                vb.write(constDeclStr);
                vb.write('\t#End Region\n\n');
            }

            if (varDelcStr.length > 0) {
                vb.write('\t#Region "Global Variables"\n');
                vb.write(varDelcStr);
                vb.write('\t#End Region\n\n');
            }

            if (propertiesStr.length > 0) {
                vb.write('\t#Region "Global Getters/Setters"\n');
                vb.write(propertiesStr);
                vb.write('\t#End Region\n\n');
            }


            if (dimModules.length > 0) {
                vb.write('\t#Region "Used Modules"\n');
                vb.write(dimModules);
                vb.write('\t#End Region\n\n');
            }

            vb.write(globalBlocks);
            vb.write('\n');

            vb.write(functionBlocksCode);

            if (aspx != null) {
                vb.write('\n\t\'************************************************');
                vb.write('\n\t\'');
                vb.write('\n\t\' Sub: Page_Load');
                vb.write('\n\t\' Sub is called on each page load');
                vb.write('\n\t\'');
                vb.write('\n\t\'************************************************');
                vb.write('\n\tProtected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load\n');
            } else {
                vb.write('\n\t\'************************************************');
                vb.write('\n\t\'');
                vb.write('\n\t\' Sub: New');
                vb.write('\n\t\' Sub is called when the class is instantiated ');
                vb.write('\n\t\'');
                vb.write('\n\t\'************************************************');
                vb.write('\n\tSub New( Page As AspPage )\n\n');
                vb.write('\t\tMyBase.New( Page )\n\n');

                if (functionMap[entry.class].vb6_Class_Initialize !== undefined) {
                    vb.write('\t\tvb6_Class_Initialize()\n\n');
                }

            }
            if (newModules.length > 0) {
                vb.write('\n');
                vb.write('\t#Region "Start Creation\n');
                vb.write(newModules);
                vb.write('\t#End Region\n');
                vb.write('\n');
            }

            remainingData = remainingData.replace(/Option\s+Explicit/gi, '').replace(/([^\n]+)/g, '\t\t$1').replace(/('.*\n(\n){2})/g, '\n').replace(/^\s*\n{2}/gm, '');

            vb.write(sanitizer.clean(remainingData) + '\n');

            vb.write('\tEnd Sub\n');

            vb.write('\n\tProtected Overrides Sub Finalize()\n');
            if (functionMap[entry.class].vb6_Class_Terminate !== undefined) {
                vb.write('\t\tvb6_Class_Terminate()\n');
            }
            vb.write('\tEnd Sub\n');

            vb.write('\n');
            vb.write('End Class\n');
        }

        if (classBlocks !== '') {
            vb.write('\n' + classBlocks + '\n');
        }
    }

    if ( aspx != null )
        aspx.end();


    if ( vb != null )
        vb.end();

    if ( writeMode && entry.data !== undefined ){
        data = vb.getContents();
        return data.toString('utf8');
    }

}

basePath = argv.base;
var startPage = argv.page;
var targetPath = argv.out;
var rabbitHoleMode = argv['rabbit-hole'];

var functionMapFile = path.join( targetPath, "function_map.json" );

// If hte function map cache exists, read it!
if ( fs.existsSync(functionMapFile) ) {
    //functionMapCache = JSON.parse(fs.readFileSync(functionMapFile));
}

sourceFiles.push( 'vb6.vb' );

sourceFiles.push( path.join( basePath, startPage ) );

console.log ("started processing at => " + new Date().toLocaleTimeString());

for( var i = 0 ; i < sourceFiles.length ; i++ ){

    //if ( i == 0 )
    //   continue;

    var file = sourceFiles[i];

    var parts = path.parse(file);
    var filePath = parts.dir;
    var fileName = parts.name;
    var subPath = path.relative( basePath, filePath );

    var outPath = path.join( targetPath, subPath.toLowerCase() );

    var className = fileName.replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');

    fileName = fileName.toLowerCase();

    var entry;

    // First file is always mine
    if ( i == 0 ){
        entry = { 'type' : 'class', 'name' : fileName, 'class' : className, 'relative' : subPath, 'in' : file, 'aspx' : null, 'vb' : path.join( targetPath, fileName + '.vb' ) };
    } else{
        var aspxFile = path.join( outPath, fileName + '.aspx' );
        var vbFile = path.join( outPath, fileName + '.aspx.vb' );

        entry = { 'type' : 'page', 'name' : fileName, 'class' : 'page' + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile };
    }

    var cacheKey = path.join( path.relative( targetPath, path.parse(entry.vb).dir ), fileName);

    if ( functionMapCache[cacheKey] != undefined ){
        functionMap = JSON.parse( JSON.stringify( functionMapCache[cacheKey] ) );
    }else{
        functionMap = {};
        // We alway include our vb6 legacy stuff
        if ( functionMapCache.vb6 != undefined )
            functionMap = functionMapCache.vb6; //JSON.parse( JSON.stringify( functionMapCache.vb6 ) );
    }

    // First process is to create a global map of functions, constants and variables for the next step of writing out the code
    processFile( entry, false, false );

    // Now write the code and handle the function mappings
    processFile( entry, rabbitHoleMode, true );

    functionMapCache[cacheKey] = functionMap;
};

fs.writeFileSync( functionMapFile, JSON.stringify( functionMapCache, null, '  ' ) );

console.log ("finished processing at => " + new Date().toLocaleTimeString());

// Write out our base class
fs.writeFileSync( path.join( targetPath, "PageClass.vb" ), fs.readFileSync('PageClass.vb') );
fs.writeFileSync( path.join( targetPath, "AspPage.vb" ), fs.readFileSync('AspPage.vb') );

writtenFiles.push(path.join( targetPath, "PageClass.vb" ));
writtenFiles.push(path.join( targetPath, "AspPage.vb" ));

var targetProjFile = path.join(targetPath, argv.project + '.vbproj');

if ( argv.overwrite || !fs.existsSync(targetProjFile) ) {

    console.log("Creating VB Project File => " + targetProjFile )
    // Write out the project
    var data = fs.readFileSync('template.vbproj');

    data = data.toString('utf8');

    data = data.replace('%GUID%', uuid.v4());

    var files = [];
    var codeFiles = [];

    for (var i = 0; i < writtenFiles.length; i++) {
        var f = writtenFiles[i];
        var parts = path.parse(f);
        var s = path.relative(targetPath, f);
        var dosPath = s.replace(/\//g, '\\');

        if (parts.ext === '.aspx') {
            files.push('<Content Include="' + dosPath + '" />');
            codeFiles.push('<Compile Include="' + dosPath + '.vb"><DependentUpon>' + dosPath + '</DependentUpon><SubType>ASPXCodeBehind</SubType></Compile>');
        }
        else {
            if (!parts.base.endsWith('.aspx.vb'))
                codeFiles.push('<Compile Include="' + dosPath + '"></Compile>');
        }

    }

    data = data.replace('%ITEMS%', files.join('\n'));
    data = data.replace('%COMPILES%', codeFiles.join('\n'));

    var proj = fs.createWriteStream(targetProjFile);

    proj.write(data);
}


console.log( 'all done' );

