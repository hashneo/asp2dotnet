var fs = require('fs');
var path = require('path');
var glob = require('glob');
var mkdirp = require('mkdirp');
var forAllAsync = exports.forAllAsync || require('forallasync').forAllAsync
    , maxCallsAtOnce = 1
var streamBuffers = require("stream-buffers");
var uuid = require('node-uuid');

var processingList = [];
var functionMap = {};   // Map of global functions which we need to tie back to the defining class later
var processedList = [];
var usageMap = {};
var functionMapCache = {};
var writtenClasses = {};

String.prototype.endsWith = function (suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
};

function sanitizeCode( code ){

    code = code.replace(/Response\.Write\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Response.Write( $1 )')
        .replace(/Response\.Redirect\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Response.Redirect( $1 )')
        .replace(/Server\.Transfer\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Server.Transfer( $1 )')
        .replace(/OTAspLogError\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'OTAspLogError( $1 )')
        .replace(/if\s+err\s+then*/gi,'if Err.Number <> 0 then')
        .replace(/isNullOrEmpty\s*(?=\(*)/gi,'String.isNullOrEmpty')
        .replace(/isEmpty\s*(?=\(*)/gi,'String.isNullOrEmpty')
        .replace(/Set\s+/gi,'')
        .replace(/\s{0}_$/gmi,' _')
        .replace(/^\s+\w*\s*sub\s+class_initialize/gmi,'Sub New')
        .replace(/^\s+\w*\s*sub\s+class_terminate/gmi,'Protected Overrides Sub Finalize')
        .replace(/Server\.CreateObject\s*\((.*)\)/gi,'System.Activator.CreateInstance(System.Type.GetTypeFromProgID( $1 ))')
        .replace(/=\s+CreateObject\s*\((.*)\)/gi,'= System.Activator.CreateInstance(System.Type.GetTypeFromProgID( $1 ))')
        .replace(/\bnull\b/gi,'DBNull.Value')
        .replace(/\barray\b\s*\((.*?)\)/gi,'New Object(){ $1 }');

    code = code.replace(/\S_$/gmi,' _');

    return code;
}

function replaceInlineCode(code) {
    var regEx = /<%[\s\r\n\t]*=\s*([\s\S]+?)[\s\r\n\t]*%>/gi;

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

function isOpenIfThenElseBlock(code) {
    var c = false;
    var regEx = /\b(if|elseif|then(?=\s*\n*)|else|end\s+if)\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word != 'end if' );
    }
    return c;
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


function processFile( entry, writeMode ) {

    var sourceFile = entry.in;

    // Prevent files being processed twice

    if ( writeMode && writtenClasses[entry.class] ){
        writtenClasses[entry.class] = true;
        console.log( 'I have already processed file => ' + sourceFile + ', skipping' );
        return;
    }
/*
    functionMap = functionMapCache[entry.class];

    if ( functionMap === undefined ){
        functionMap = {};
    }
*/

    var codeBlocks = [];
    var functionBlocks = [];
    var includeFiles = [];

    var data = fs.readFileSync(sourceFile);

    if ( !writeMode ){
        console.log( 'pre-reading and merging file => ' + sourceFile );
    }

    var sourcePath = path.dirname( sourceFile );
    var targetPath = path.dirname(  entry.vb );


    data = data.toString('utf8');

    data = data.replace(/_\r\n(\t*|\s*)/gi,'');

    var match;

    // If there is no ASP VBScript tag, we move to module mode and don't create an ASP file
    if ( !data.match( /<%@.*%>/g ) ){
        entry.aspx = null;
        entry.vb = entry.vb.replace('inc_', '').replace('.aspx.vb', '.vb');
        var className = entry.name.toLowerCase().replace('inc_','').replace('const_','').replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');
        entry.class = 'cls' + className;
    }

    mkdirp.sync(targetPath, function(err) {
        console.log('could not create dir => ' + targetPath );
        throw err;
    });

    // Strip out includes
    var inclRegex = /<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->/gi;

    var processingIncludeList = [];

    var htmlIncludes = [];

    while (( match = inclRegex.exec(data) ) != null) {

        var matchFile = match[2].replace(/\\/g,'/');

        var includeFile = includeFile = path.normalize(path.join( sourcePath ,matchFile )).trim();

        if ( !fs.existsSync(includeFile) ){
            console.log( 'WARNING: include file => ' + includeFile + ' does not exists, skipping' );
            continue;
        }

        var file = includeFile;

        var parts = path.parse(file);
        var filePath = parts.dir;
        var fileName = parts.name;

        if ( parts.ext.toLowerCase() !== '.asp' ) {
            htmlIncludes.push( {'file': matchFile, 'tag': match[0]} );
            continue;
        }

        var subPath = path.relative( sourcePath, filePath );

        var outPath = path.normalize( path.join( targetPath, subPath ) );

        var className = fileName.toLowerCase().replace('inc_','').replace('const_','').replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');

        var aspxFile = path.join( outPath, className + '.aspx' );
        var vbFile = path.join( outPath, className + '.aspx.vb' );

        includeFiles.push( { 'file' : includeFile, 'class' : 'cls' + className } );

        processFile( { 'name' : fileName, 'class' : 'cls' + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile }, writeMode );
    }

    // Inject whatever html the file has included
    for ( var i = 0 ; i < htmlIncludes.length ; i++ ){
        var htmlInclude = htmlIncludes[i];
        //var contents = fs.readFileSync(htmlInclude.file).toString();
        var contents = '<% Response.WriteFile ("' + htmlInclude.file+ '")" %>';
        data = data.replace( htmlInclude.tag, contents );
    }

    // Process the include files first as we need a function map later

    var aspx = null;

    if ( writeMode ){
        if ( entry.aspx != null ){
            console.log( 'writing aspx source file => ' + entry.aspx );
            aspx = fs.createWriteStream(entry.aspx);
        }

        var vb = fs.createWriteStream(entry.vb);

        console.log( 'writing vb source file => ' + entry.vb );
    }

    var os = new streamBuffers.WritableStreamBuffer({
        initialSize: (100 * 1024),      // start as 100 kilobytes.
        incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
    });

    // Remove classes
    var regEx = /((?!'(?:\n|\r)|(?:\n|\r))(?:\s*'.*?\r\n)*(class\s*(\w+))(?:[\s\S]+?)(?:end\s+(?:class)))/gmi;

    var remainingData = data;//.toString('utf8');

    var classBlocks = '';
    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[1];
        var className = match[3];
        classBlocks += sanitizeCode(codeBlock) +'\n';
        remainingData = remainingData.replace(codeBlock, "");
    }

    // Strip out code blocks
    var codeRegEx = /(?:<%|<SCRIPT\s+LANGUAGE\s*=\s*"VBScript"\s+RUNAT\s*=\s*"Server">)[\s\r\n\t]*(?=[^=|@])([\s\S]+?)[\s\r\n\t]*(?:%>\s*|<\/SCRIPT>(?!.*")\s*)/gi;

    data = remainingData;
    while (( match = codeRegEx.exec(data) ) != null) {
        var codeBlock = match[1];

        codeBlocks.push({'code': codeBlock, 'start': match.index, 'length': match[0].length});
    }

    // If we have pure HTML code then we need to create a couple of dummy blocks
    // at the start and the end of the file so the rest of the code will function
    if (codeBlocks.length == 0) {
        var start = 0;
        if (match = /(<%@.*%>.*\r\n)/g.exec(data)) {
            start = match.index + match[0].length;
        }

        codeBlocks.push({'code': '', 'start': start, 'length': 0});
        codeBlocks.push({'code': '', 'start': data.length, 'length': 0});
    }

    // Iterate over code blocks looking for inline HTML
    var isInSub = false;

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
        if (isOpenIfThenElseBlock(code) ||
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
        htmlChunk = htmlChunk.replace(/<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->\r\n*/g, '');

        if (aspx != null && writeMode)
            aspx.write(replaceInlineCode(htmlChunk));

        os.write('\n');
        os.write(sanitizeCode(thisBlock.code));

        if (thisBlock.write !== undefined) {
            var lines = thisBlock.code.split(/\r\n|\r|\n/);
            var m = lines[lines.length - 1].match(/(\s+).*/);
            var indent = '';
            if (m != null && m.length > 1)
                indent = m[1];

            var htmlCode = thisBlock.write.replace(/"/g, '""'); //.replace(/\t|\r|\n/gi, '');
            var regEx = /([^\r\n]+)/gi;
            while (( match = regEx.exec(htmlCode) ) != null) {
                var line = sanitizeCode( replaceInlineCode(match[1]) );

                if ( line.length > 0 )
                    os.write('\n' + indent + 'Response.Write ("' +  line + '")');
            }
        }

        os.write('\n');

        if (nextBlock !== null) {
            srcPos = nextBlock.start;
        } else {
            srcPos = endPos;
        }
    }

    // Strip out code functions and subs
    var regEx = /((?!'(?:\n|\r)|(?:\n|\r))(?:\s*'.*?\r\n)*(?:public|private\s+|)(?:sub\s+(\w+)\s*\(*.*\)*|function\s+(\w+)\s*\(.*\))(?:[\s\S]+?)(?:end\s+(?:sub|function))($:\r\n)*)/gi;

    // Get the stream data and convert it to a string
    data = os.getContents();

    remainingData = data.toString('utf8');
    data = remainingData;

    functionMap[entry.class] = {};  // Create a map of functions for later use

    var functionBlocks = '';

    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[1];

        var fnName = match[2];
        if (fnName === undefined)
            fnName = match[3];

        if (fnName !== undefined) {
            functionMap[entry.class][fnName] = {'name': fnName, 'hits': 0};
        }

        functionBlocks += codeBlock.replace(/([^\r\n]+)/g, '\t$1') + '\n\n';

        remainingData = remainingData.replace(codeBlock, "");
    }

    function parseConst(code) {
        var comments = '';

        if (match = /((?:\s*'.*?\r\n)+)(?=const\s+)/gi.exec(code)) {
            comments = match[1];
            code = code.replace(match[1], '');
        }

        if (match = /const\s+(\w+)\s*=\s*([\S\w]+|".*")*\s*('.*)*/gi.exec(code)) {
            var result = {
                'name': match[1],
                'var': match[1],
                'value': match[2],
                'comment': ( match[3] != undefined ? match[3] : comments ).replace(/^'/gm, '').replace(/\r/g, '').split('\n'),
                'hits': 0
            };
        }
        else {
            throw "unable to parse const => " + code;
        }

        return result;
    }

    function parseDim(code) {
        var comments = '';

        if (match = /((?:\s*'.*?\r\n)+)(?=dim\s+)/gi.exec(code)) {
            comments = match[1].replace(/^'/gm, '').replace(/\r/g, '').split('\n');
            code = code.replace(match[1], '');
        }

        if (match = /('.*)/gi.exec(code)) {
            comments = match[1].replace(/^'/gm, '').replace(/\r/g, '').split('\n'),
                code = code.replace(match[1], '');
        }

        var definitions = code.split(',');

        var results = [];
        for (var i = 0; i < definitions.length; i++) {

            var def = definitions[i].replace(/^dim\s+/i, '');
            if (match = /(\w+)\s*:*\s*(\w+\s*=\s*(.*|".*"))*\s*('.*)*/gi.exec(def)) {
                var thisVar = match[1].replace('g_', '').replace(/^x+_/gmi, '').replace(/_x+$/gmi, '').replace(/_/g, ' ').replace(/(\b[a-z](?!\s))/g, function (x) {
                    return x.toUpperCase();
                }).replace(/ /g, '');
                var result = {
                    'name': match[1],
                    'var': thisVar,
                    'init': match[2],
                    'value': match[3],
                    'comment': ( i == 0 ? comments : undefined ),
                    'hits': 0
                };

                results.push(result);
            } else {
                throw "unable to parse dim => " + definitions[i];
            }
        }

        return results;
    }

    // Strip out all DIM and CONST declarations
    regEx = /^(((?:'.*?\r\n){0,}?)(?:public\s+)?(const|dim)+\s+(?:[\s\S]+?))$/gmi;

    data = remainingData;

    var globalBlocks = '';
    var constDecls = [];
    var variableDecls = [];

    while (( match = regEx.exec(data) ) != null) {
        var codeBlock = match[1];//.replace( /\r?\n|\r/g, '' );
        remainingData = remainingData.replace(match[0], "");

        switch (match[3].toLowerCase()) {
            case 'dim':
                variableDecls.push(parseDim(codeBlock));
                break;
            case 'const':
                constDecls.push(parseConst(codeBlock));
                break;
            default:
                globalBlocks += '\t' + codeBlock;
        }
    }

    functionMap[entry.class]['_Variables'] = variableDecls;
    functionMap[entry.class]['_Constants'] = constDecls;

    var constDeclStr = '';

    // Only do subsitutions when we are writing
    if ( writeMode ) {

        // Remove all extra comments and multiple new lines
        remainingData = remainingData.replace(/('.*\r\n|(\r\n){2})/g, '\n').replace(/^\s*\n{2}/gm, '');

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
            var Line = '\tPublic Const ' + constDecls[i].var;
            Line += new Array(Math.max(0, 50 - Line.length)).join(' ') + ' = ' + constDecls[i].value;
            if (constDecls[i].comment !== undefined) {
                Line = formatCommentBlock(constDecls[i].comment) + Line;
            }
            constDeclStr += Line + '\n';
        }

        var varDelcStr = '';

        for (var i = 0; i < variableDecls.length; i++) {
            for (var x = 0; x < variableDecls[i].length; x++) {
                var Line = '\tPublic Property ' + variableDecls[i][x].var;
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

                var varName = varName = "_" + cls.substring(3);

                var inUse = false;

                function doSubstitution(_type, _what, _with) {

                    var regEx
                    /*
                     if ( _type === 'var')
                     regEx = new RegExp('^(?:(?!sub|function).)*\\s*[&=+]+\\s*[^a-z.](' +  _what.name + ')[^a-z.]|[^a-z.](' +  _what.name + ')[^a-z.]\\s*[&=+]?.*$', 'gmi')
                     else if ( _type === 'const')
                     regEx = new RegExp('^(?:(?!sub|function).)*\\s*[&=+]+\\s*[^a-z.](' + _what.name + ')[^a-z.].*$', 'gmi')
                     else
                     */
                    //regEx = new RegExp('^(?:(?!sub|function).)*\\s*[&=+]*\\s*[^a-z.](' + _what.name + ')[^a-z.](\\s+|\\({1}|\\n+).*$', 'gmi')
                    regEx = new RegExp('^(?:(?!sub|function|' + _with + ').)*\\s*\\W(\\b' + _what.name + '\\b)\\W.*$', 'gmi');

                    //  ^(?:(?!sub|function).)*\s*(?:[&=+]+\s*[^a-z.](nothing)[^a-z.]|[^a-z.](nothing)[^\.]\s*[&=+]?).*$

                    code = code.replace( regEx, function(match, p1, p2, p3, offset, string){
                        _what.hits += 1;
                        inUse = true;
                        return match.replace( new RegExp('\\W' + p1 + '\\W', 'gi'), function(m)
                        {
                            return m[0] + _with + m[m.length-1];
                        });
                    });
                }

                for (var name in functionMap[cls]) {
                    if (name === '_Variables') {
                        for (var i = 0; i < functionMap[cls][name].length; i++) {
                            for (var x = 0; x < functionMap[cls][name][i].length; x++) {
                                var _with = varName + '.' + functionMap[cls][name][i][x].var;
                                if (cls === entry.class)
                                    _with = '_' + functionMap[cls][name][i][x].var;
                                doSubstitution('var', functionMap[cls][name][i][x], _with);
                            }
                        }
                    } else if (name === '_Constants') {
                        for (var i = 0; i < functionMap[cls][name].length; i++) {
                            var _with = varName + '.' + functionMap[cls][name][i].var;
                            if (cls === entry.class)
                                _with = functionMap[cls][name][i].var;
                            doSubstitution('const', functionMap[cls][name][i], _with);
                        }
                    } else {
                        if (cls !== entry.class)
                            doSubstitution('method', functionMap[cls][name], varName + '.' + name);
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
                    //break;
                }
            }

            return code;//.replace(/aAbBcC_|_aAbBcC/g, '');
        }

        functionBlocks = processFunctionMap(functionBlocks);

        remainingData = processFunctionMap(remainingData);

        var dimModules = ''
        var newModules = ''

        for (var i = 0; i < usedModules.length; i++) {
            dimModules += '\tDim ' + usedModules[i].var + '\n';
            newModules += '\t\t' + usedModules[i].var + ' = new ' + usedModules[i].class + '( Server, Application, Request, Response )\n';
        }


        if (aspx != null) {
            vb.write('Public Class ' + entry.class + '\n\n');
            vb.write('\tInherits Page' + '\n');
            vb.write('\n');
        } else {
            vb.write('Public Class ' + entry.class + '\n\n');
            vb.write('\tDim Server\n');
            vb.write('\tDim Application\n');
            vb.write('\tDim Request\n');
            vb.write('\tDim Response\n');
            vb.write('\n');
        }
        if (constDeclStr.length > 0) {
            vb.write('\t\'--------- Start Globals Constants ---------\n');
            vb.write(constDeclStr);
            vb.write('\t\'--------- End Globals Constants ---------\n\n');
        }

        if (varDelcStr.length > 0) {
            vb.write('\t\'--------- Start Globals Variables ---------\n');
            vb.write(varDelcStr);
            vb.write('\t\'--------- End Globals Variables ---------\n\n');
        }

        if (dimModules.length > 0) {
            vb.write('\t\'--------- Start Modules used ---------\n');
            vb.write(dimModules);
            vb.write('\t\'--------- End Modules used ---------\n\n');
        }

        vb.write(globalBlocks);
        vb.write('\n');

        vb.write(functionBlocks);

        if (aspx != null) {
            vb.write('\n\tProtected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load\n');
        } else {
            vb.write('\n\tSub New( Server, Application, Request, Response )\n\n');
            vb.write('\t\tMe.Server = Server\n');
            vb.write('\t\tMe.Application = Application\n');
            vb.write('\t\tMe.Request = Request\n');
            vb.write('\t\tMe.Response = Response\n');
        }
        if (newModules.length > 0) {
            vb.write('\n');
            vb.write('\t\'--------- Start Module Creation ---------\n');
            vb.write(newModules);
            vb.write('\t\'--------- End Module Creation ---------\n');
            vb.write('\n');
        }
        vb.write(remainingData.replace(/Option\s+Explicit/gi, '').replace(/([^\r\n]+)/g, '\t\t$1').replace(/('.*\r\n|(\r\n){2})/g, '\n').replace(/^\s*\n{2}/gm, '') + '\n');

        vb.write('\tEnd Sub\n');
        vb.write('\n');
        vb.write('End Class\n');

        if (classBlocks !== '') {
            vb.write('\n' + classBlocks + '\n');
        }
    }

    functionMapCache[entry.class] = functionMap;
}

var searchString = process.argv[2];
var sourcePath = path.parse( searchString );

if ( sourcePath.ext === '' )
    searchString += '/**/*.asp';

var targetPath = process.argv[3];

glob( searchString, function( err, files ) {

    for( var i = 0 ; i < files.length ; i++ ){

        var file = files[i];

        var parts = path.parse(file);
        var filePath = parts.dir;
        var fileName = parts.name;
        var subPath = path.relative( sourcePath.dir, filePath );

        var outPath = path.join( targetPath, subPath );

        var className = fileName.replace(/_/g,' ').replace(/(\b[a-z](?!\s))/g, function(x){ return x.toUpperCase();}).replace(/ /g,'');

        var aspxFile = path.join( outPath, className + '.aspx' );
        var vbFile = path.join( outPath, className + '.aspx.vb' );

        functionMap = {};

        // First process is to create a global map of functions, constants and variables for the next step of writing out the code
        processFile( { 'name' : fileName, 'class' : 'page' + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile }, false );

        // Now write the code and handle the function mappings
        processFile( { 'name' : fileName, 'class' : 'page' + className, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile }, true );

    }

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

        console.log( 'all done' );
    });
*/
});


