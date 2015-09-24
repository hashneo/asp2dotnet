var fs = require('fs');
var path = require('path-extra');
var mkdirp = require('mkdirp');
var streamBuffers = require("stream-buffers");
var glob = require('glob');

var uuid = require('node-uuid');
var isNumeric = require("isnumeric")

var argv = require('yargs')
    .usage('Usage: $0 <options>')
    .options({
        'base': {
            demand: true,
            describe: 'Base path of ASP Site (Virtual Dir in IIS)',
            type: 'string'
        },
        'page':{
            demand: true,
            describe: 'ASP page(s) to process. Wildcards are also ok. !!Beware of Globbing!!. Use /blah/\\*\\*/*.asp for recursion.',
            type: 'array'
        },
        'out' : {
            demand: true,
            describe: 'Base output path to write files to. Any include writes will be relative to this path (beware of ../../path/script.asp!)',
            type: 'string'
        },
        'namespace' : {
            demand: true,
            describe: 'VB.Net project namespace',
            type: 'string'
        },
        'project' : {
            demand: true,
            describe: 'VB.Net project to create. Relative to output path.',
            type: 'string'
        },
        'overwrite' : {
            demand: false,
            describe: 'Overwrite all files, ignoring timestamps and File Protected: directive',
            type: 'boolean'
        },
        'no-rename' : {
            demand: false,
            describe: 'Prevent renaming of global variables into nicer looking ones',
            type: 'boolean'
        },
        'no-includes' : {
            demand: false,
            describe: 'Only process the given ASP files, ignore includes',
            type: 'boolean'
        },
        'rabbit-hole' : {
            demand: false,
            describe: 'Look for linked ASP files and add them to the list of files to be processed',
            type: 'boolean'
        },
        'verbose' : {
            demand: false,
            describe: 'Print a whole lot of information about what is going on to console',
            type: 'boolean'
        }
    })
    .wrap(null)
    .version(function() {
        return require('./package').version;
    })
    .argv;


var variablesParser = require('./variables-parser');
var propertiesParser = require('./properties-parser');
var functionsParser = require('./functions-parser');
var sanitizer = require('./code-sanitizer');


require('./string-extensions.js');

var functionMap = {};   // Map of global functions which we need to tie back to the defining class later
var functionMapCache = {};
var writtenClasses = {};
var basePath;
var sourceFiles = [];
var writtenFiles = [];
var targetNamespace;

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
    var sourceStat = sourceFile != undefined ? fs.statSync(sourceFile) : undefined;

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

    var origHeader;
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
        data = data.replace(/_\s*\n(\t*|\s*)/gi,'').replace(/<%\s+=/gi,'<%=');

        // fix up ampersands that have no space between tem and words
        //data = data.replace(/&"/gi,'& "').replace(/"&/gi,'" &');
    }

    var match;

    if ( entry.data === undefined ) {

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

            var includeFile = path.normalize(path.join(sourcePath, matchFile)).trim();

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
                    'level' : entry.level + 1,
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
            if ( !argv.overwrite && fs.existsSync(entry.aspx) ){
                if ( argv.verbose )
                    console.log( 'INFO: aspx file already exists => ' + entry.aspx + ', skipping' );
            }else{
                console.log( 'Writing aspx source file => ' + entry.aspx );
                aspx = fs.createWriteStream(entry.aspx);
            }
            writtenFiles.push(entry.aspx);
        }

        if ( entry.vb != null )
            writtenFiles.push(entry.vb);

        var converterHeader;

        if ( fs.existsSync(entry.vb) ){
            var fd = fs.openSync(entry.vb, 'r');
            var buffer = new Buffer(1024);
            fs.readSync(fd, buffer, 0, buffer.length, 0);
            var fileHeader = buffer.toString('utf8');
            fs.closeSync(fd);

            var match;

            if ( ( match = /#Region\s+"asp2dotnet\s+converter\s+header"([\s\S]+?)#End Region/gi.exec(fileHeader) ) != null ){
                fileHeader = match[1].trim();
                var regEx =  /^'\s+(.*?):\s+(.*)$/gmi;
                while ( ( match = regEx.exec( fileHeader ) ) != null ){
                    if ( converterHeader === undefined )
                        converterHeader = {};
                    converterHeader[match[1]] = match[2];
                }
            }
        }

        if ( entry.vb != null ) {
            var fileProtected = converterHeader != undefined && converterHeader['File Protected'] !== undefined ? JSON.parse(converterHeader['File Protected']) : false;
            var origModified = converterHeader != undefined && converterHeader['Original Modified'] !== undefined ? Date.parse(converterHeader['Original Modified']) : null;
            var hasChanged = origModified !== undefined ? +origModified != +sourceStat.mtime : true;

            var skipFile = false;

            if ( !argv.overwrite ) {
                if (fileProtected && argv.verbose)
                    console.log('INFO: vb file => ' + entry.vb + ' has File Protected = true in the header, skipping');

                if (!hasChanged)
                    console.log('Source file => ' + sourceFile + ' has not changed, skipping');

                skipFile = fileProtected | !hasChanged;
            }

            if (!skipFile) {
                console.log('Writing vb source file => ' + entry.vb);
                // var _fd = fs.openSync( entry.vb, 'w' );
                // vb = fs.createWriteStream( '', {fd:_fd} );
                vb = fs.createWriteStream( entry.vb );
            }
        }else{
            if ( argv.verbose )
                console.log('INFO: writing class => ' + entry.class + ' to a temp stream');
            vb = new streamBuffers.WritableStreamBuffer({
                initialSize: (100 * 1024),      // start as 100 kilobytes.
                incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
            });
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

        if ( entry.name !== 'vb6'){
            classData = processFile( { 'type' : 'class', 'level' : entry.level, 'name' : fileName, 'class' : className,  'data' : '<%\n\n' + classData + '\n\n%>' }, rabbitHoleMode, writeMode );
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
        origHeader = p1;
        return '';
    });

    // First remove all Functions
    var result = functionsParser.parse( remainingData, argv.verbose, entry.in );
    var functionBlocks = result.blocks;
    functionMap[entry.class] = result.map;
    functionMap[entry.class]['_Level'] = entry.level;
    functionMap[entry.class]['_Type'] = entry.type;

    // Next we remove all Properties (GET/LET)
    var result = propertiesParser.parse( result.data, argv.verbose );
    functionMap[entry.class]['_Properties'] = result.properties;

    var globalBlocks = '';

    // Finally get all of the Variables
    result = variablesParser.parse( result.data, argv.verbose, entry.type === 'class' || (argv.rename == false) );

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

                for( var functionName in functionMap[entry.class]){

                    if ( functionName[0] === '_')
                        continue;
                    if( functionName.equalsIgnoreCase(varDecl[x].var) ){
                        varDecl[x].var = '_' + varDecl[x].var;
                        /*
                        varDecl[x].visibility = 'Protected';

                        for ( var propName in functionMap[entry.class]['_Properties'] ){
                            if ( propName.equalsIgnoreCase( varDecl[x].var ) ){
                                break;
                            }
                        }*/
                    }9
                }

                for ( var propName in functionMap[entry.class]['_Properties'] ){
                    if ( propName.equalsIgnoreCase( varDecl[x].var ) ){
                        varDecl[x].var = '_' + varDecl[x].var;
                    }
                }

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
                if ( cls === entry.class || functionMap[cls]._Type == 'page' )
                    continue;

               // if ( functionMap[cls]['_Level'] > entry.level )
               //     continue;

                includeClasses.push( cls );
            }

            for (var a = 0 ; a < includeClasses.length ; a++ ){

                var cls = includeClasses[a];

                var varName = "_" + cls; //.substring(3);

                var inUse = false;

                function isReservedName(n){
                    n = n.toLowerCase();

                    if ( n === 'new') return true;
                    if ( n === '_type') return true;
                    if ( n === '_level') return true;

                    return false;
                }

                if ( functionMap[cls]['_Type'] === 'class'){
                    code = code.substitute( { 'name' : 'new\\s+' + cls }, 'New ' + cls + '( Page )', function(match){
                        if ( match.indexOf('.') > 0 ){
                            var whichClass = match.split('.')[0];
                            if ( !whichClass.equalsIgnoreCase( item.name) )
                                return false;
                        }
                        //inUse = true;
                        return true;
                    });

                    // We won't resolve classes as globals since you should have
                    // already thought about it. Our Vb6 class is a special case!
                    if ( cls !== 'Vb6' )// && cls !== entry.class)
                        continue;

                }

                for (var name in functionMap[cls]) {

                    if ( isReservedName(name) || ( thisFunction !== undefined && name.toLowerCase() === thisFunction.name.toLowerCase() ) )
                        continue;

                    if (name[0] === '_') {

                        functionMap[cls][name].forEach( function( item ){

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

                            code = code.substitute(item, _with, function(match){

                                if ( match.indexOf('.') > 0 ){
                                    var whichClass = match.split('.')[0];
                                    if ( !whichClass.equalsIgnoreCase( item.name) )
                                        return false;
                                }
                                inUse = true;
                                return true;
                            });
                        });

                    } else {

                        if ( localVariables !== undefined && localVariables.contains(name) )
                            continue;

                        var _with = varName + '.' + name;
                        if (cls === entry.class)
                            _with = 'Me.' + name;
                        if ( functionMap[cls][name].type === 'class')
                            _with = cls + '.' + name
                        if ( functionMap[cls][name].global )
                            _with = cls + '.' + name
                        if ( functionMap[cls][name].visibility.toLowerCase() === 'private')
                            continue;
                        code = code.substitute( functionMap[cls][name], _with, function(match){

                            if ( match.indexOf('.') > 0 ){
                                var whichClass = match.split('.')[0];
                                if ( !whichClass.equalsIgnoreCase( functionMap[cls][name].name) )
                                    return false;
                            }
                            inUse = true;
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

            var getStr = '';
            var setStr = '';

            if ( classProperty.get !== undefined ){
                var theGet = classProperty.get;

                getStr = (theGet.visibility.toLowerCase() !== 'public' ? theGet.visibility : '') + ' Get\n' +
                    processFunctionMap(theGet) +
                    'End Get\n'
            }

            if ( classProperty.let !== undefined ){
                var theSet = classProperty.let;
                setStr = (theSet.visibility.toLowerCase() !== 'public' ? theSet.visibility : '') + ' Set(' +  theSet.setParam + ' As Object )\n' +
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
            functionBlocksCode += functionBlock.comment +
                                  functionBlock.function.signature + '\n' +
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
            aspx.write( '<%@ Page Language="VB" AutoEventWireup="true" CodeBehind="' + entry.name + '.aspx.vb" Inherits="' + targetNamespace + '.' + entry.class + '" %>' )
        }

        if ( sourceFile !== undefined ) {

            vb.write('#Region \"asp2dotnet converter header\"\n');
            vb.write('\' Source file: "file://' + sourceFile + '"\n');
            vb.write('\' Original Modified: ' + sourceStat.mtime.toISOString() + '\n');
            vb.write('\' Date Converted: ' + new Date().toISOString() + '\n');
            vb.write('\' File Protected: false\n');
            vb.write('#End Region\n');
            vb.write('\n');
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

            if (origHeader !== undefined) {
                vb.write('\t#Region "Original Header"\n');
                vb.write(origHeader + '\n');
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


    if ( vb != null ) {
        vb.end();
        //fs.closeSync(vb.fd);
    }

    if ( writeMode && entry.data !== undefined ){
        data = vb.getContents();

      //  if ( entry.type === 'class')
      //      functionMap[entry.class] == undefined;
        return data.toString('utf8');
    }

}

basePath = argv.base;
var targetPath = argv.out;
var rabbitHoleMode = argv['rabbit-hole'];

targetNamespace = argv.namespace;

var functionMapFile = path.join( targetPath, "function_map.json" );

// If hte function map cache exists, read it!
if ( fs.existsSync(functionMapFile) ) {
    //functionMapCache = JSON.parse(fs.readFileSync(functionMapFile));
}

sourceFiles.push( 'vb6.vb' );

for( var i in argv.page ){
    sourceFiles =  sourceFiles.concat( glob.sync( path.join( basePath, argv.page[i] ) ) );
}

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

    entry['level'] = 0;

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

//fs.writeFileSync( functionMapFile, JSON.stringify( functionMapCache, null, '  ' ) );

console.log ("finished processing at => " + new Date().toLocaleTimeString());

// Write out our base class
fs.writeFileSync( path.join( targetPath, "PageClass.vb" ), fs.readFileSync('PageClass.vb') );
fs.writeFileSync( path.join( targetPath, "AspPage.vb" ), fs.readFileSync('AspPage.vb') );
fs.writeFileSync( path.join( targetPath, "web.config" ), fs.readFileSync('web.config') );

writtenFiles.push(path.join( targetPath, "PageClass.vb" ));
writtenFiles.push(path.join( targetPath, "AspPage.vb" ));


var targetProjFile = path.resolve( path.join( targetPath, argv.project ) );
var targetProjectParts = path.parse(targetProjFile);
var targetProjectPath = targetProjectParts.dir;

targetProjectParts.ext = '.vbproj';
targetProjectParts.base = targetProjectParts.name + targetProjectParts.ext;

targetProjFile = path.format(targetProjectParts);

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
        var s = path.relative(targetProjectPath, f);
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
    data = data.replace('%NAMESPACE%', targetNamespace );

    var proj = fs.createWriteStream(targetProjFile);

    proj.write(data);
}

console.log( 'all done' );

