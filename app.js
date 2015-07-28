var fs = require('fs');
var path = require('path');
var mkdirp = require('mkdirp');
var streamBuffers = require("stream-buffers");
var argv = require('optimist').argv;
var uuid = require('node-uuid');
var isNumeric = require("isnumeric");

var functionMap = {};   // Map of global functions which we need to tie back to the defining class later
var functionMapCache = {};
var writtenClasses = {};
var basePath;
var sourceFiles = [];

String.prototype.endsWith = function (suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
};

String.prototype.splice = function (pos,size) {
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

function sanitizeCode( code ){

    code = code.replace(/Response\.Write\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Response.Write( $1 )$2')
        .replace(/Response\.Redirect\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Response.Redirect( $1 )$2')
        .replace(/Server\.Transfer\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Server.Transfer( $1 )$2')
        .replace(/if\s+err\s+then*/gi,'if Err.Number <> 0 Then')
        .replace(/(\s+)set(\s+)/gi,'$1$2')
        .replace(/(\s+)wend(\s+)/gi,'$1End While$2')
        .replace(/\s{0}_$/gmi,' _')
        .replace(/^\s+\w*\s*sub\s+class_initialize/gmi,'Sub New')
        .replace(/^\s+\w*\s*sub\s+class_terminate/gmi,'Protected Overrides Sub Finalize')
        .replace(/Server\.CreateObject\b\s*\((.*)\)/gi,'System.Activator.CreateInstance(System.Type.GetTypeFromProgID( $1 ))')
        .replace(/=\s*CreateObject\b\s*\((.*)\)/gi,'= System.Activator.CreateInstance(System.Type.GetTypeFromProgID( $1 ))')
        .replace(/(\s+|\b)null\b(\s*)/gi,'$1DBNull.Value$2')
        //.replace(/(\s+|\b)Not\s+IsObject\s*\((.*)\)/gi,'$1$2 Is Nothing ')
        //.replace(/(\s+|\b)IsObject\s*\((.*)\)/gi,'$1Not ($2 Is Nothing) ')
        //.replace(/(\s+|\b)IsNull\s*\((.*)\)/gi,'$1($2 = DBNull.Value) ')
        .replace(/(\s+|\b)Empty\b(\s*)/gi,'$1Nothing$2')
        .replace(/(\s+|\b)Timer\b\s*\(\s*\)?(\s*)/gi,'$1New Timer()$2')
        .replace(/(\s+|\b)Date\b\s*\(\s*\)?(\s*)/gi,'$1New Date()$2')
        .replace(/Request\.QueryString\b(\s*\(.*?\))?(.*)/gi,'Request.QueryString$1.ToString()$2')
        .replace(/Request\.Cookies\b(\s*\(.*?\))?(.*)/gi,'Request.Cookies$1.ToString()$2')
        .replace(/err\.raise\b(\s+\w+.*)/gi,'Err.Raise($1)');


    // Nice replacements
    code = code.replace(/\.append\s+(.+?)(?='|$)/gmi,'.Append($1)');

        //.replace(/Append\s+(.*)/gi, 'Append( $1 ) ')

    //code = code.functionReplace( 'isobject', '($1 Is Nothing)' );

    code = code.functionReplace( 'round', 'Math.Round' );
    code = code.functionReplace( 'array', 'New Object()', '{', '}' );

    code = code.replace(/\s+isNullOrEmpty\s*(?=\(*)/gi, function(match, p1){
        return match;
    });

    code = code.replace(/\s+isEmpty\s*(?=\(*)/gi, function(match, p1){
        return ' isNullOrEmpty ';
    });

    //.replace(,'String.isNullOrEmpty')
    //.replace(,'String.isNullOrEmpty')

    code = code.replace(/\S_$/gmi,' _');

    return code;
}

function replaceInlineCode(code) {
    var regEx = /<%[\s\n\t]*=\s*([\s\S]+?)[\s\n\t]*%>/gi;

    var copy =  code;
    while (( match = regEx.exec(code) ) != null) {
        var func = match[1];
        func = func.replace(/""/g, '"');
        copy = copy.replace(match[1], func);
    }

    code = copy.replace(regEx, '" & $1 & "');

    return code;
}

function isOpenSelectCase(code) {
    var c = false;
    var regEx = /\b(select\s+case(?=\s*\n*)|case(?=\s*.*\n*)|end\s+select)\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word != 'end select' );
    }
    return c;
}

function isOpenIfThenElseBlock(code, nestCounter) {
    var regEx = /\b(if|elseif|then(?=\s*\n*)|else|end\s+if)\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        switch ( word ){
            case 'if':
                break;
            case 'elseif':
                break;
            case 'then':
                nestCounter.if++;
                break;
            case 'else':
                break;
            case 'end if':
                if ( nestCounter.if > 0 )
                    nestCounter.if--;
                break;
        }
    }
    return nestCounter.if > 0;
}


function isOpenDoLoop(code) {
    var c = false;
    var regEx = /\b(do(?=\s+)(while|until)*|while|loop(?=\s+)(while|until)*|exit\s+do|wend)\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word != 'loop' && word != 'wend' );
    }
    return c;
}

function isOpenForLoop(code) {
    var c = false;
    var regEx = /\b((for\s+)|exit\s+for|next)\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word != 'next' );
    }
    return c;
}


function isOpenSubFunction(code) {
    var c = false;
    var regEx = /^\s*\b((sub|function)(?=\(*.*\)*)|(end\s+(sub|function)))\b/gmi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        if ( word == 'sub' || word == 'function' )
            c = true;
        else if  ( word == 'end sub' || word == 'end function' )
            c = false;
    }
    return c;
}

function isEndSubFunction(code) {
    var c = false;
    var regEx = /^\s*\b((sub|function)(?=\(*.*\)*)|(end\s+(sub|function)))\b/gmi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word == 'end sub' || word == 'end function' );
        break;
    }
    return c;
}


function processFile( entry, rabbitHoleMode, writeMode ) {

    var sourceFile = entry.in;

    var vbSourceFile = path.parse(sourceFile).ext.toLowerCase() === '.vb';

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
    var functionBlocks = [];
    var includeFiles = [];

    var data = fs.readFileSync(sourceFile);

    if ( !writeMode ){
        if ( argv.verbose )
            console.log( 'pre-reading and merging file => ' + sourceFile );
    }

    var sourcePath = path.dirname( sourceFile );
    var targetPath = path.dirname(  entry.vb );

    data = data.toString('utf8')

    // Firstly replace all CRLF with LF (Windows)
    data = data.replace(/\r\n/g,'\n');

    // Replace any remaining CR with LF (Mac OS 9)
    data = data.replace(/\r/g,'\n');

    // Now clean the rest since we are In Unix Mode. Get rid of VBScript _ line separators, fix string concats, code markers
    data = data.replace(/_\s*\n(\t*|\s*)/gi,'').replace(/&"/gi,'& "').replace(/"&/gi,'" &').replace(/<%\s+=/gi,'<%=');

    var match;

    // If there is no ASP VBScript tag, we move to module mode and don't create an ASP file
    if ( !data.match( /<%@.*%>/g ) ){
        entry.aspx = null;
        entry.vb = entry.vb.replace('.aspx.vb', '.vb');
        var className = entry.name.toLowerCase().replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');

        var prefix = 'cls';
        entry.class = prefix + className;
    }

    mkdirp.sync(targetPath, function(err) {
        console.log('could not create dir => ' + targetPath );
        throw err;
    });

    var regEx;

    // Strip out includes
    regEx = /<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->/gi;

    var htmlIncludes = [];

    data = data.replace( regEx, function( match, p1, p2 ){

        var matchFile = p2.replace(/\\/g,'/');

        var includeFile = includeFile = path.normalize(path.join( sourcePath ,matchFile )).trim();

        if ( !fs.existsSync(includeFile) ){
            console.log( 'WARNING: include file => ' + includeFile + ' does not exists, skipping' );
            return '';
        }

        var file = includeFile;

        var parts = path.parse(file);
        var filePath = parts.dir;
        var fileName = parts.name;

        if ( parts.ext.toLowerCase() !== '.asp' ) {
            htmlIncludes.push( {'file': matchFile, 'tag': match[0]} );
            return '';
        }

        var subPath = path.relative( sourcePath, filePath).toLowerCase();

        var outPath = path.normalize( path.join( targetPath, subPath ) );

        var className = fileName.replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');

        var prefix = 'cls';
        fileName = fileName.toLowerCase();
        var aspxFile = path.join( outPath, fileName + '.aspx' );
        var vbFile = path.join( outPath, fileName + '.aspx.vb' );

        includeFiles.push( { 'file' : includeFile, 'class' : prefix + className } );

        if ( argv.includes !== false )
            processFile( { 'name' : fileName, 'class' : prefix + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile }, rabbitHoleMode, writeMode );

        return '';
    });


    // Inject whatever html the file has included
    for ( var i = 0 ; i < htmlIncludes.length ; i++ ){
        var htmlInclude = htmlIncludes[i];
        var contents = '<% Response.WriteFile ("' + htmlInclude.file+ '")" %>';
        data = data.replace( htmlInclude.tag, contents );
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
            }
        }

        if ( argv.overwrite == undefined && fs.existsSync(entry.vb) ){
            console.log( 'WARNING: vb file already exists => ' + entry.vb + ', skipping' );
        }else{

            var skipFile = false;
            if ( fs.existsSync(entry.vb) ){
                var fd = fs.openSync(entry.vb, 'r');
                var buffer = new Buffer(256);
                fs.readSync(fd, buffer, 0, buffer.length, 0);
                skipFile = (buffer.toString('utf8').indexOf('\'ignore') == 0);
                fs.closeSync(fd);
            }

            if (skipFile){
                console.log( '\'ignore declaration found in vb source file => ' + entry.vb );
            }else{
                vb = fs.createWriteStream(entry.vb);
                console.log( 'writing vb source file => ' + entry.vb );
            }
        }
    }

    var os = new streamBuffers.WritableStreamBuffer({
        initialSize: (100 * 1024),      // start as 100 kilobytes.
        incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
    });

    // Look for Linked ASP files as we will add them to the list we need to process
    if (!vbSourceFile){

        var regEx = /"(([-a-zA-Z0-9/\-_]+?\.asp)(?!\w+)(\?*.*)?)"/gi;
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
    regEx = /((?!'(?:\n)|(?:\n))(?:\s*'.*?(?:\n))*(?:public|private)?\s*\b(class\b\s*(\w+)\s*(?:'.*)?(?:\n))(?:[\s\S]*?)(?:end\s+(?:class)))/gmi;

    var remainingData = data;

    var classBlocks = '';
    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[1];
        var className = match[3];
        classBlocks += sanitizeCode( codeBlock ) +'\n';
        remainingData = remainingData.replace(codeBlock, "");
    }

    // Strip out code blocks
   //var codeRegEx = /(?:<%|<SCRIPT\s+LANGUAGE\s*=\s*"VBScript"\s+RUNAT\s*=\s*"Server">)[\s\n\t]*(?=[^=|@])([\s\S]+?)[\s\n\t]*(?:%>\s*|<\/SCRIPT>(?!.*")\s*)/gi;


    if (!vbSourceFile) {

        // Replace RUNAT Server Tag with VBScript Marker
        regEx = /<SCRIPT\s+LANGUAGE\s*=\s*"VBScript"\s+RUNAT\s*=\s*"Server">([\s\n\t]*(?=[^=|@])([\s\S]+)[\s\n\t]*)<\/SCRIPT>/gi;

        remainingData = remainingData.replace(regEx, function (match, p1) {
            return '<%\n' + p1 + '%>\n';
        });

        var codeRegEx = /(?:<%)[\s\n\t]*(?=[^=|@])([\s\S]+?)[\s\n\t]*(?:%>\s*)/gi;

        data = remainingData;
        while (( match = codeRegEx.exec(data) ) != null) {
            var codeBlock = match[1];

            codeBlocks.push({'code': codeBlock, 'start': match.index, 'length': match[0].length});
        }

        // If we have pure HTML code then we need to create a couple of dummy blocks
        // at the start and the end of the file so the rest of the code will function
        if (codeBlocks.length == 0) {
            var start = 0;
            if (match = /(<%@.*%>.*\n)/g.exec(data)) {
                start = match.index + match[0].length;
            }

            codeBlocks.push({'code': '', 'start': start, 'length': 0});
            codeBlocks.push({'code': '', 'start': data.length, 'length': 0});
        }

        // Iterate over code blocks looking for inline HTML
        var isInSub = false;

        var nestCounter = {'if': 0};

        // If its less than 2 blocks, no way can it have HTML in between
        for (var i = 0; i < codeBlocks.length - 1; i++) {
            var thisBlock = codeBlocks[i];
            var nextBlock = codeBlocks[i + 1];

            if (i == 6) {
                i = 6;
            }

            // Get code - the comments
            var code = thisBlock.code.replace(/\s*'.*/gi, '');

            if (isInSub && isEndSubFunction(code)) {
                isInSub = false;
            }

            if (!isInSub && isOpenSubFunction(code)) {
                isInSub = true;
            }

            // If we read any open open ended statements, capture the code as we will
            // convert to a Response.Write
            if (isOpenIfThenElseBlock(code, nestCounter) ||
                isOpenDoLoop(code) ||
                isOpenForLoop(code) ||
                isOpenSelectCase(code) || isInSub) {
                var startPos = thisBlock.start + thisBlock.length;
                var endPos = nextBlock.start;
                var htmlChunk = data.slice(startPos, endPos);
                thisBlock['write'] = htmlChunk;//.toString('utf8');
            }

        }

        var srcPos = 0;

        for (var i = 0; i < codeBlocks.length; i++) {
            var thisBlock = codeBlocks[i];
            var nextBlock = null;

            var endPos = thisBlock.start + thisBlock.length;

            if (thisBlock.write !== undefined) {
                nextBlock = codeBlocks[i + 1];
            }

            var htmlChunk = data.slice(srcPos, thisBlock.start);

            if (codeBlocks.length == 1) {
                htmlChunk = data;
            }

            htmlChunk = htmlChunk.replace(/<%@.*%>/g, '<%@ Page Language="VB" AutoEventWireup="true" CodeBehind="' + entry.name + '.aspx.vb" Inherits="' + entry.class + '" %>');

            if (aspx != null && writeMode)
                aspx.write(replaceInlineCode(htmlChunk));

            os.write('\n');
            os.write( sanitizeCode( thisBlock.code ) );

            if (thisBlock.write !== undefined) {
                var lines = thisBlock.code.split(/\n/);
                var m = lines[lines.length - 1].match(/(\s+).*/);
                var indent = '';
                if (m != null && m.length > 1)
                    indent = m[1];

                var htmlCode = thisBlock.write.replace(/"/g, '""'); //.replace(/\t||\n/gi, '');
                var regEx = /([^\n]+)/gi;

                while (( match = regEx.exec(htmlCode) ) != null) {
                    var line = sanitizeCode( replaceInlineCode(match[1]) );

                    if (line.length > 0)
                        os.write('\n' + indent + 'Response.Write ("' + line + '")');
                }
            }

            os.write('\n');

            if (nextBlock !== null) {
                srcPos = nextBlock.start;
            } else {
                srcPos = endPos;
            }
        }
        // Get the stream data and convert it to a string
        data = os.getContents();
        remainingData = data.toString('utf8');
    }

    // Strip out code functions and subs
    var regEx =/^\s*((?:(public|private)\s+)?((?:sub|function)\s+(\w+)(\s*\(\s*(.*?)\s*\))?){1}(?:[\s\S]+?)(?:end\s+(?:sub|function)){1})[^\n]*$/gmi;

    data = remainingData;

    data = data.replace( /('\s*\$Header[\s\S]+?All\sRights\sReserved.)/gi, function(m,p1){
        fileHeader = p1;
        return '';
    });

    remainingData = data;

    functionMap[entry.class] = {};  // Create a map of functions for later use

    var functionBlocks = '';
    var totalRemoved = 0;
    while (( match = regEx.exec(data) ) != null) {

        var codeBlock = match[1];
        var visibility = match[2];
        var fnSignature = match[3];
        var fnName = match[4];
        var parameters = match[6];

        if ( visibility === undefined )
            visibility = "Public";

        if ( parameters === undefined )
            parameters = '';

        var endPos = match.index + match[0].length + 1;

        if ( argv.verbose ){
            console.log("Found Function = > " + fnSignature );
        }

        var commentBlock = '';

        var i = match.index + match[0].indexOf(fnSignature);
        var previousLine = '';
        var maxBlankLines = 1;
        var startPos = i
        while (--i > 0) {
            if (data[i] == '\n') {
                if (previousLine.length > 0) {
                    if (previousLine[0] == '\'' || previousLine[0] == '<')
                        commentBlock = previousLine + '\n' + commentBlock;
                    else {
                        if ( commentBlock.length > 0 ){
                            // If we ate a blank line adjust the startpos to consume the character
                            startPos -= (1 - (maxBlankLines + 1));
                        }
                        break;
                    }
                }else{
                    if ( commentBlock.length == 0 && maxBlankLines-- < -1 ){
                        break;
                    }
                }
                previousLine = '';
            } else {
                previousLine = data[i] + previousLine;
            }
        }

        startPos -= commentBlock.length;

        codeBlock = commentBlock + codeBlock;

        if (fnName !== undefined) {
            functionMap[entry.class][fnName] =
            {   'name': fnName,
                'visibility' : visibility,
                'signature' : fnSignature,
                'parameters' : parameters.trim().length > 0 ? parameters.split(',') : undefined,
                'hits': 0
            };
        }

        functionBlocks += codeBlock.replace(/([^\n]+)/g, '\t$1') + '\n\n';

        // Remove the comments and codeblock after the first \n
        var startStr = remainingData.substr(0,startPos - totalRemoved);
        var endStr = remainingData.substr(endPos - totalRemoved);

        totalRemoved += (endPos - startPos);

        remainingData = startStr + endStr;
        //remainingData = remainingData.replace( codeBlock, '' );

        if ( remainingData.indexOf(fnSignature) != -1 ) {
            console.log( 'processing error => ' + entry.in);
            console.log( fnSignature + ", still exists in the remaining data, aborting!" );
            process.exit(1);
        }
    }

    function parseConst(code) {
        var comments = '';

        if (match = /((?:\s*'.*?\n)+)(?=\bconst\b\s+)/gi.exec(code)) {
            comments = match[1];
            code = code.replace(match[1], '');
        }

        // Sanity check to make sure we don't have any commented out variables
        if ( code.search(/^\s*'.*?\bdim\b.*$/gmi) != -1 )
            return null;

        if (match = /(private|public)?\s*const\s+(\w+)(?:\s*as\s*(\w+))?\s*=\s*(".*"|[\S\w]+)*\s*('.*)*/gi.exec(code)) {

            var visibility = match[1];

            if ( visibility === undefined )
                visibility = "Public";

            var result = {
                'visibility' : visibility,
                'name': match[2],
                'var': match[2],
                'type': match[3],
                'value': match[4],
                'comment': ( match[5] != undefined ? match[5] : comments ).replace(/^'/gm, '').split('\n'),
                'hits': 0
            };

            // Prevent short consts like 'rs' making into the wild
            if ( result.name.length < 4 ){
                if ( writeMode )
                    console.log( 'WARNING: Const => ' + result.name + ' is less than 4 chars, marking it Private!');
                result.visibility = 'Private';
            }

        }
        else {
            throw "unable to parse const => " + code;
        }

        return result;
    }

    function parseDim(code) {
        var comments;

        var start = code.search(/^\s*((public|private)\s+)?\bdim\b/mi);

        comments = code.substr( 0, start );
        code = code.substr( start );

        // Sanity check to make sure we don't have any commented out variables
         if ( code.search(/^\s*'.*?\bdim\b.*$/gmi) != -1 )
           return null;

        if ( comments.trim().length > 0 ){
            comments = comments.replace(/^'/gm, '').split('\n');
            // If there are 2 previous blank lines before the comments, assume its junk and scrap the comment
            if ( comments.length > 2 && comments[comments.length-1].trim().length == 0 &&  comments[comments.length-2].trim().length == 0 )
                comments = [];
        }else
            comments = [];


        if (match = /('.*)/gi.exec(code)) {
            comments = match[1].replace(/^'/gm, '').split('\n');
                code = code.replace(match[1], '');
        }

        code = code.replace(/[\n]/g,'');

        var results = [];

        var visibility;

        if (match = /(\s*(?:(public|private)?\s*dim)\s*)/gi.exec(code)) {
            code = code.replace(match[1], '');
            visibility = match[2];
        }

        if ( visibility === undefined )
            visibility = "Public";

        code = code.replace(/^dim\s+/i, '');
        var regEx = /(\w+(?:\s*\(.*?\))*)\s*:*\s*(\w+\s*=\s*(.*|".*"))*\s*('.*)*/gi;

        while (( match = regEx.exec(code)) != null) {
            var thisVar = match[1].replace('g_', '').replace(/^x+_/gmi, '').replace(/_x+$/gmi, '').replace(/_/g, ' ').replace(/(\b[a-z](?!\s))/g, function (x) {
                return x.toUpperCase();
            }).replace(/ /g, '');

            var result = {
                'visibility': visibility,
                'name': match[1],
                'var': thisVar,
                'init': match[2],
                'value': match[3],
                'comment': ( results.length == 0 ? comments : undefined ),
                'hits': 0
            };

            // Prevent short variables like 'rs' making into the wild
            if ( result.name.length < 4 ){
                if (isNumeric(entry.name)){
                    entry.name = '_' + entry.name;
                    if ( writeMode )
                        console.log( 'BIG WARNING: Variable => ' + result.name + ' has a numeric name!. I\'ll rename with and underscore to prevent mass updates');
                }
                if ( writeMode )
                    console.log( 'WARNING: Variable => ' + result.name + ' is less than 4 chars, marking it Private!');
                result.visibility = 'Private';
            }


            results.push(result);
        }

        if ( results.length == 0 ) {
            throw "unable to parse dim => " + code;
        }

        return results;
    }

    // Strip out all DIM and CONST declarations
    regEx = /^(((?:'.*?\n){0,}?)(?:private|public)?\s*\b(const|dim)+\b\s+(?:[\s\S]+?))\s*$/gmi;

    data = remainingData;

    var globalBlocks = '';
    var constDecls = [];
    var variableDecls = [];

    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[0];

        var endPos = match[0].length;
        if ( codeBlock.replace('','').trim().endsWith(',')){
            codeBlock += '\n';
            var offset = match.index + match[0].length + 1;
            var pos;
            while ( (pos = data.indexOf('\n', offset)) != -1){
                var nextLine = data.substr( offset, pos - offset );
                offset += nextLine.length + 1;
                endPos = offset;
                //nextLine = nextLine.replace('','');
                codeBlock += nextLine + '\n';
                if ( !nextLine.replace('','').trim().endsWith(',') )
                    break;
            }
        }

        remainingData = remainingData.replace(codeBlock, "");

        switch (match[3].toLowerCase()) {
            case 'dim':
                var variableDecl = parseDim(codeBlock);
                if (variableDecl != null )
                    variableDecls.push(variableDecl);
                else
                    globalBlocks += '\t' + codeBlock;
                break;
            case 'const':
                var constDecl = parseConst(codeBlock);
                if (constDecl != null)
                    constDecls.push(constDecl);
                else
                    globalBlocks += '\t' + codeBlock;
                break;
            default:
                globalBlocks += '\t' + codeBlock;
        }
    }

    functionMap[entry.class]['_Variables'] = variableDecls;
    functionMap[entry.class]['_Constants'] = constDecls;

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

        for (var i = 0; i < constDecls.length; i++) {
            var type = constDecls[i].type;
            if ( type === undefined ) type = "Object"
            var Line = '\t' + constDecls[i].visibility + ' Const ' + constDecls[i].var + ' As ' + type;
            Line += new Array(Math.max(0, 50 - Line.length)).join(' ') + ' = ' + constDecls[i].value;
            if (constDecls[i].comment !== undefined) {
                Line = formatCommentBlock(constDecls[i].comment) + Line;
            }
            constDeclStr += Line + '\n';
        }

        var varDelcStr = '';

        for (var i = 0; i < variableDecls.length; i++) {
            for (var x = 0; x < variableDecls[i].length; x++) {
                var type = variableDecls[i][x].type;
                if ( type === undefined ) type = "Object"
                var Line = '\t' + variableDecls[i][x].visibility + ' Property ' + variableDecls[i][x].var  + ' As ' + type;
                if (variableDecls[i][x].value !== undefined)
                    Line += new Array(Math.max(0, 50 - Line.length)).join(' ') + ' = ' + variableDecls[i][x].value;
                if (variableDecls[i][x].comment !== undefined) {
                    Line = formatCommentBlock(variableDecls[i][x].comment) + Line;
                }
                varDelcStr += Line + '\n';
            }
        }

        var usedModules = [];

        function processFunctionMap(code) {

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

                var varName = "_" + cls.substring(3);

                var inUse = false;

                function doSubstitution(_type, _what, _with) {

                    var regEx

                    //regEx = new RegExp('((?:sub|function)\\s*\\w+\\s*\\(.*|' + _with.split('.')[0] + '\\.)?\\s*(\\b' + _what.name + '\\b)', 'gmi');

                    regEx = new RegExp('^\\n?\\s*((?=(?:public|private)?\\s*(sub|function))|.*)(\\b' + _what.name + '\\b).*$', 'gmi' )

                    code = code.replace( regEx, function(m, p1, p2, p3, offset, string){

                        // I have no idea why the above regex is ignoring
                        if ( p1.match(/\b(sub|function)\b/gi) != null && p2 === undefined ){
                            return m;
                        }

                        if (p2 !== undefined){
                            return m;
                        }

                        p1 = p3;

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
                            // Are we in a comment. if there are quotes also in there, ignore as the comment block could be
                            // buried in the quotes
                            //if ( ( m.match(new RegExp( '\'.*\\b' + _what.name + '\\b.*','gi' ))) != null && m.indexOf('"') == -1 )
                            //   return m;
                        }

                        if ( m.indexOf('"')!= -1 ){
                            var r = new RegExp( '(["\'])(?:\\b(' + _what.name + ')\\b.*?)?\\1','gi' );
                            if ( ( (m2 = m.match(r))) != null ) {
                                if ( m2[0].indexOf(_what.name) != -1)
                                    return m;
                            }
                        }


                        _what.hits += 1;
                        inUse = true;

                        if ( _what.parameters !== undefined && _what.parameters.length > 0 ){
                            if (  m.match( new RegExp( p1 + '\\s*\\(') ) == null ){
                                m = m.replace( new RegExp('^(.*)' + p1 + '\\s+(?!\\s*=)(.+?)(?=(?:$))', 'm'), function(m2, p2, p3){
                                    return p2 + p1 + '( ' + p3 + ' )';
                                });
                            }
                        }
                        return m.replace( new RegExp('\\b' + p1 + '\\b', 'gi'), function(m2){
                            return  _with;
                        });
                    });
                }

                for (var name in functionMap[cls]) {

                    if (name === '_Variables') {
                        for (var i = 0; i < functionMap[cls][name].length; i++) {
                            for (var x = 0; x < functionMap[cls][name][i].length; x++) {
                                var _with = varName + '.' + functionMap[cls][name][i][x].var;
                                if (cls === entry.class)
                                    _with = functionMap[cls][name][i][x].var;
                                if ( functionMap[cls][name][i][x].visibility.toLowerCase() === 'private')
                                    continue;
                                doSubstitution('var', functionMap[cls][name][i][x], _with);
                            }
                        }
                    } else if (name === '_Constants') {
                        for (var i = 0; i < functionMap[cls][name].length; i++) {
                            var _with = cls + '.' + functionMap[cls][name][i].var;
                            if ( functionMap[cls][name][i].visibility.toLowerCase() === 'private')
                                continue;
                            if (cls === entry.class)
                                _with = functionMap[cls][name][i].var;
                            doSubstitution('const', functionMap[cls][name][i], _with);
                        }
                    } else {
                        var _with = varName + '.' + name;
                        if (cls === entry.class)
                            _with = name;
                        if ( functionMap[cls][name].type === 'class')
                            _with = cls + '.' + name
                        if ( functionMap[cls][name].visibility.toLowerCase() === 'private')
                            continue;
                        doSubstitution('function', functionMap[cls][name], _with);
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

        if ( !vbSourceFile ) {
            functionBlocks = processFunctionMap(functionBlocks);
            remainingData = processFunctionMap(remainingData);
        }

        var dimModules = ''
        var newModules = ''

        for (var i = 0; i < usedModules.length; i++) {
            dimModules += '\tPrivate ' + usedModules[i].var + ' As ' + usedModules[i].class + '\n';
            newModules += '\t\t' + usedModules[i].var + ' = new ' + usedModules[i].class + '( ' + (aspx != null ? 'Me' : 'Me.Page' ) + ' )\n';
        }

        if (aspx != null) {
            vb.write('Public Class ' + entry.class + '\n\n');
            vb.write('\tInherits Page' + '\n');
            vb.write('\n');
        } else {
            vb.write('Public Class ' + entry.class + '\n\n');
            vb.write('\tInherits PageClass' + '\n')

            vb.write('\n');
        }

        if ( fileHeader !== undefined){
            vb.write('\t\'--------- Start Original Header ---------\n');
            vb.write( fileHeader + '\n' );
            vb.write('\t\'--------- End Original Header ---------\n\n');
        }

        if (constDeclStr.length > 0) {
            vb.write('\t\'--------- Start Global Constants ---------\n');
            vb.write(constDeclStr);
            vb.write('\t\'--------- End Global Constants ---------\n\n');
        }

        if (varDelcStr.length > 0) {
            vb.write('\t\'--------- Start Global Variables ---------\n');
            vb.write(varDelcStr);
            vb.write('\t\'--------- End Global Variables ---------\n\n');
        }

        if (dimModules.length > 0) {
            vb.write('\t\'--------- Start Used Modules ---------\n');
            vb.write(dimModules);
            vb.write('\t\'--------- End Used Modules ---------\n\n');
        }

        vb.write(globalBlocks);
        vb.write('\n');

        vb.write(functionBlocks);

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
            vb.write('\n\tSub New( Page As System.Web.UI.Page )\n\n');
            vb.write('\t\tMyBase.New( Page )\n\n');
            /*
            vb.write('\t\t\' Store our page variables fore use in the class\n' );
            vb.write('\t\tMe.Server = Server\n');
            vb.write('\t\tMe.Application = Application\n');
            vb.write('\t\tMe.Request = Request\n');
            vb.write('\t\tMe.Response = Response\n\n');
            */
        }
        if (newModules.length > 0) {
            vb.write('\n');
            vb.write('\t\'--------- Start Module Creation ---------\n');
            vb.write(newModules);
            vb.write('\t\'--------- End Module Creation ---------\n');
            vb.write('\n');
        }

        remainingData = remainingData.replace(/Option\s+Explicit/gi, '').replace(/([^\n]+)/g, '\t\t$1').replace(/('.*\n(\n){2})/g, '\n').replace(/^\s*\n{2}/gm, '');

        vb.write( remainingData + '\n');

        vb.write('\tEnd Sub\n');
        vb.write('\n');
        vb.write('End Class\n');

        if (classBlocks !== '') {
            vb.write('\n' + classBlocks + '\n');
        }
    }

    if ( aspx != null )
        aspx.end();

    if ( vb != null )
       vb.end();

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
        entry = { 'name' : fileName, 'class' : className, 'relative' : subPath, 'in' : file, 'aspx' : null, 'vb' : path.join( targetPath, fileName + '.vb' ) };
    } else{
        var aspxFile = path.join( outPath, fileName + '.aspx' );
        var vbFile = path.join( outPath, fileName + '.aspx.vb' );

        entry = { 'name' : fileName, 'class' : 'page' + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile };
    }

    var cacheKey = path.join( path.relative( targetPath, path.parse(entry.vb).dir ), fileName);

    if ( functionMapCache[cacheKey] != undefined ){
        functionMap = JSON.parse( JSON.stringify( functionMapCache[cacheKey] ) );
    }else{
        functionMap = {};
        // We alway include our vb6 legacy stuff
        if ( functionMapCache.vb6 != undefined )
            functionMap = JSON.parse( JSON.stringify( functionMapCache.vb6 ) );
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

/*
    fs.readFile('template.vbproj', function (err, data) {

        data = data.toString('utf8');

        data = data.replace('%GUID%', uuid.v4() );

        var files = [];
        var codeFiles = [];

        for ( var i = 0 ; i < processingList.length ; i++ ) {
            var f = processingList[i];
            var dosPath = f.relative.replace(/\//g,'\\')  + '\\';
            files.push( '<Content Include="' + dosPath + f.name + '.aspx" />' );
            codeFiles.push( '<Compile Include="' + dosPath + f.name + '.aspx.vb">\n<DependentUpon>'+ f.name + '.aspx</DependentUpon>\n<SubType>ASPXCodeBehind</SubType>\</Compile>' );
        }

        data = data.replace('%ITEMS%', files.join('\n'));
        data = data.replace('%COMPILES%', codeFiles.join('\n'));

        var proj = fs.createWriteStream(path.join( targetPath, 'project.vbproj' ) );

        proj.write( data );

    });
*/

console.log( 'all done' );

