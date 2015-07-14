var fs = require('fs');
var path = require('path');
var glob = require('glob');
var mkdirp = require('mkdirp');
var forAllAsync = exports.forAllAsync || require('forallasync').forAllAsync
    , maxCallsAtOnce = 1
var streamBuffers = require("stream-buffers");
var uuid = require('node-uuid');

var processingList = [];
var functionMap = {};
var processedList = [];

function sanitizeCode( code ){

    code = code.replace(/Response\.Write\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Response.Write( $1 )')
        .replace(/Response\.Redirect\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Response.Redirect( $1 )')
        .replace(/Server\.Transfer\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'Server.Transfer( $1 )')
        .replace(/OTAspLogError\s+(.*(\s+&\s+vbCrLf)?)\n*/gi,'OTAspLogError( $1 )')
        .replace(/if\s+err\s+then*/gi,'if Err.Number <> 0 then')
        .replace(/Server\.CreateObject\((.*)\)/gi,'System.Activator.CreateInstance(System.Type.GetTypeFromProgID( $1 ))');

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
    var regEx = /\b((sub|function)(?=\(*.*\)*)|(end\s+(sub|function)))\b/gi;
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
    var regEx = /\b((sub|function)(?=\(*.*\)*)|(end\s+(sub|function)))\b/gi;
    while (( match = regEx.exec(code) ) != null) {
        var word = match[1].toLowerCase();
        c = ( word == 'end sub' || word == 'end function' );
        break;
    }
    return c;
}


function processFile( complete, entry, i ) {

    var sourceFile = entry.in;
    var aspxFile = entry.aspx;
    var vbFile = entry.vb;

    // Prevent files being processed twice
    for ( var i = 0 ; i < processedList.length ; i++ ){
        if ( ( processedList[i].toLowerCase() === sourceFile.toLowerCase() ) ){
             complete();
             return;
         }
    }

    processedList.push( sourceFile );

    var sourcePath = path.dirname( sourceFile );
    var targetPath = path.dirname( vbFile );

    if ( entry.name.indexOf('inc_') == 0 ){
        entry.name = entry.name.replace('inc_', '');
        aspxFile = null;
        vbFile = vbFile.replace('inc_', '').replace('.aspx.vb', '.vb');
    }

    entry.name = '_' + entry.name;

    String.prototype.endsWith = function (suffix) {
        return this.indexOf(suffix, this.length - suffix.length) !== -1;
    };

    var codeBlocks = [];
    var functionBlocks = [];
    var includeFiles = [];

    mkdirp.sync(targetPath, function(err) {
        console.log('could not create dir => ' + targetPath );
        throw err;
    });

    var aspx = null;

    if ( aspxFile != null )
        aspx = fs.createWriteStream(aspxFile);

    var vb = fs.createWriteStream(vbFile);

    var os = new streamBuffers.WritableStreamBuffer({
        initialSize: (100 * 1024),      // start as 100 kilobytes.
        incrementAmount: (10 * 1024)    // grow by 10 kilobytes each time buffer overflows.
    });

    fs.readFile(sourceFile, function (err, data) {
        if (err) throw err;

        var match;

        // Strip out includes
        var inclRegex = /<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->/gi;

        var processingIncludeList = [];
        while (( match = inclRegex.exec(data) ) != null) {
            var includeFile = path.normalize(path.join(sourcePath, match[2])).trim();

            includeFiles.push( includeFile );
            var file = includeFile;

            var parts = path.parse(file);
            var filePath = parts.dir;
            var fileName = parts.name;
            var subPath = path.relative( sourcePath, filePath );

            var outPath = path.normalize( path.join( targetPath, subPath ) );
            var aspxFile = path.join( outPath, fileName + '.aspx' );
            var vbFile = path.join( outPath, fileName + '.aspx.vb' );

           processingIncludeList.push( { 'name' : fileName, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile } );
        }

        // Process the include files first as we need a function map later

        forAllAsync(processingIncludeList, processFile, maxCallsAtOnce).then(function () {

            console.log( 'processing file => ' + sourceFile );

            // Strip out code blocks
            var codeRegEx = /(?:<%|<SCRIPT\s+LANGUAGE\s*\=\s*"VBScript"\s+RUNAT\s*=\s*"Server">)[\s\r\n\t]*(?=[^=|@])([\s\S]+?)[\s\r\n\t]*(?:%>|<\/SCRIPT>)/gi;


            // Remove classes
            var regEx = /^((\s+class\s*(\w+))(?:[\s\S]+?)(?:end\s+(?:class))$)/gmi;

            var remainingData = data.toString('utf8');

            var classBlocks = '';
            while (( match = regEx.exec(data) ) != null) {
                var codeBlock = match[1];
                var className = match[3];

                classBlocks += codeBlock;

                remainingData = remainingData.replace( codeBlock, "" );
            }

            data = remainingData;

            while (( match = codeRegEx.exec(data) ) != null) {
                var codeBlock = match[1];

                codeBlocks.push({'code': codeBlock, 'start': match.index, 'length': match[0].length});
            }

            var isInSub = false;

            for (var i = 0; i < codeBlocks.length - 1; i++) {
                var thisBlock = codeBlocks[i];
                var nextBlock = codeBlocks[i + 1];

                if ( i == 6 ){
                    i = 6;
                }

                // Get code - the comments
                var code = thisBlock.code.replace(/\s*'.*/gi, '');
                ;

                if ( isInSub && isEndSubFunction(code) ){
                    isInSub = false;
                }

                if ( !isInSub && isOpenSubFunction(code) ) {
                    isInSub = true;
                }


                if (isOpenIfThenElseBlock(code) || isOpenDoLoop(code) || isOpenForLoop(code) || isOpenSelectCase(code) || isInSub ) {

                    var startPos = thisBlock.start + thisBlock.length;
                    var endPos = nextBlock.start;
                    var htmlChunk = data.slice(startPos, endPos);
                    thisBlock['write'] = htmlChunk.toString('utf8');
                }else{
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

                var htmlChunk = data.slice(srcPos, thisBlock.start).toString('utf8');

                htmlChunk = htmlChunk.replace( /<%@.*%>/g, '<%@ Page Language="VB" AutoEventWireup="true" CodeBehind="' +  entry.name + '.aspx.vb" Inherits="_' +  entry.name + '" %>' );
                htmlChunk = htmlChunk.replace( /<!--\s*\#include\s+(file|virtual)\s*="(.*)"\s*-->\r\n*/g, '' );

                if ( aspx != null )
                    aspx.write(replaceInlineCode(htmlChunk));

                os.write('\n');
                os.write( sanitizeCode( thisBlock.code ) );

                if (thisBlock.write !== undefined) {
                    var lines = thisBlock.code.split(/\r\n|\r|\n/);
                    var m = lines[ lines.length - 1].match(/(\s+).*/);
                    var indent = '';
                    if (m != null && m.length > 1 )
                        indent = m[1];

                    var htmlCode = thisBlock.write.replace(/"/g, '""'); //.replace(/\t|\r|\n/gi, '');
                    var regEx = /([^\r\n]+)/gi;
                    while (( match = regEx.exec(htmlCode) ) != null) {
                        var line = match[1];
                        os.write('\n' + indent + 'Response.Write ("' + replaceInlineCode(line) + '")');
                    }
                }

                os.write('\n');

                if (nextBlock !== null) {
                    srcPos = nextBlock.start;
                } else {
                    srcPos = endPos;
                }
            }

            data = os.getContents();

            remainingData = data.toString('utf8');

            if ( entry.name === 'OpenAdStream'){
                var a = 1;
            }

            // Strip out code functions and subs

            data = remainingData;

            //var remainingData = data.toString('utf8');

            var regEx = /((?!'(?:\n|\r)|(?:\n|\r))(?:\s*'.*?\r\n)*(?:sub\s*(\w+)\s*\(*.*\)*|function\s*(\w+)\s*\(.*\))(?:[\s\S]+?)(?:end\s+(?:sub|function))(\r|\n)*)/gi;

            var functionBlocks = '';
            while (( match = regEx.exec(data) ) != null) {
                var codeBlock = match[1];

                var fnName = match[2];
                if ( fnName === undefined )
                    fnName = match[3];

                if ( fnName !== undefined ){
                    if ( functionMap[fnName] !== undefined ){
                        console.log( 'WARNING : ' + fnName + ' was previously defined in ' + functionMap[fnName].class + ', now defined in ' +  entry.name, '. This one will be ignored!' );
                    }
                    else{
                        functionMap[fnName] = { 'class' : entry.name };
                    }
                }

                functionBlocks += codeBlock.replace( /([^\r\n]+)/g, '\t$1' );

                remainingData = remainingData.replace( codeBlock, "" );
            }

            regEx = /((?!'\n|\n)(?:\s*'.*?\n)*(?:const|dim)+(?:[\s\S]+?))\n/gi;

            data = remainingData;

            var globalBlocks = '';

            while (( match = regEx.exec(data) ) != null) {
                var codeBlock = match[1];

                globalBlocks += codeBlock.replace( /([^\r\n]+)/g, '\t$1' );

                remainingData = remainingData.replace( codeBlock, "" );
            }

            if ( aspx != null ) {
                vb.write('Public Class _' + entry.name + '\n\n');
                vb.write('\tInherits Page' + '\n');
                vb.write('\n');
            }else{
                vb.write('Public Class ' + entry.name + '\n\n');
                vb.write('\tdim Server\n');
                vb.write('\tdim Application\n');
                vb.write('\tdim Request\n');
                vb.write('\tdim Response\n');
                vb.write('\n');
            }

            vb.write( globalBlocks );
            vb.write( '\n' );

            for ( var name in functionMap){
                var fClass = functionMap[name].class;
                if ( fClass !== entry.name ){

                    if ( entry.name == 'MANDATORY' && name == 'GetDBConnectionString' ){
                        var a = 1;
                    }
                    regEx = new RegExp( '(?=\\W*)(' + name + ')(?=\\W+)', 'g' );
                    functionBlocks = functionBlocks.replace( regEx, fClass + '.' + name );
                }
            }

            vb.write( functionBlocks );

            if ( aspx != null ){
                vb.write( '\n\tProtected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load\n');
            }else{
                vb.write( '\n\tSub New( Server, Application, Request, Response )\n');
                vb.write( '\t\tMe.Server = Server\n');
                vb.write( '\t\tMe.Application = Application\n');
                vb.write( '\t\tMe.Request = Request\n');
                vb.write( '\t\tMe.Response = Response\n');
            }

            vb.write( remainingData.replace( /([^\r\n]+)/g, '\t\t$1').replace('Option Explicit','') + '\n');
            vb.write( '\tEnd Sub\n');
            vb.write( '\n' );
            vb.write( 'End Class\n' );

            if ( classBlocks !== '' ){
                vb.write( '\n' + classBlocks + '\n' );
            }

            complete();
        });

    });
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
        var aspxFile = path.join( outPath, fileName + '.aspx' );
        var vbFile = path.join( outPath, fileName + '.aspx.vb' );

        processingList.push( { 'name' : fileName, 'relative' : subPath, 'in' : file, 'aspx' : aspxFile, 'vb' : vbFile } );
    }

    forAllAsync(processingList, processFile, maxCallsAtOnce).then(function () {

        fs.readFile('template.vbproj', function (err, data) {

            data = data.toString('utf8');

            data = data.replace('%GUID%', uuid.v4() );

            var files = [];
            var codeFiles = [];
            /*
            for ( var i = 0 ; i < processingList.length ; i++ ) {
                var f = processingList[i];
                var dosPath = f.relative.replace(/\//g,'\\')  + '\\';
                files.push( '<Content Include="' + dosPath + f.name + '.aspx" />' );
                codeFiles.push( '<Compile Include="' + dosPath + f.name + '.aspx.vb">\n<DependentUpon>'+ f.name + '.aspx</DependentUpon>\n<SubType>ASPXCodeBehind</SubType>\</Compile>' );
            }

            data = data.replace('%ITEMS%', files.join('\n'));
            data = data.replace('%COMPILES%', codeFiles.join('\n'));
             */
            var proj = fs.createWriteStream(path.join( targetPath, 'project.vbproj' ) );

            proj.write( data );

            console.log( 'all done' );
        });


    });

});


