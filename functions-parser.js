FunctionsParser = function(){

    require('./string-extensions.js');

    var sanitizer = require('./code-sanitizer');
    var variablesParser = require('./variables-parser');

    this.parse = function(data, verbose, srcFile){

        // Strip out code functions and subs
        var regEx =/^\s*(((?:(public|protected|private)\s+)?(shared\s+)?(?:(?:overrides)\s+)?(sub|function)\s+(\w+)(\s*\(\s*(.*?)\s*\))?(?:\s+as\s\w+)?){1}([\s\S]+?)(?:end\s+(?:sub|function)){1})[^\n]*$/gmi;
        var match;

        var remainingData = data;

        var functionMap = {};  // Create a map of functions for later use

        var functionBlocks = [];

        var totalRemoved = 0;
        while (( match = regEx.exec(data) ) != null) {

            var fnSignature = match[2];
            var visibility = match[3];
            var global = match[4] === undefined ? false : true ;
            var type = match[5];
            var fnName = match[6];
            var parameters = match[8];
            var codeBlock = match[9];

            if ( visibility === undefined )
                visibility = "Public";

            if ( parameters === undefined )
                parameters = '';

            var endPos = match.index + match[0].length + 1;

            if ( verbose ){
                console.log("Found Function = > " + fnSignature );
            }

            var commentBlock = '';

            var offset = match[0].indexOf(fnSignature);

            var i = match.index + offset;
            var previousLine = '';
            var maxBlankLines = 1;
            var startPos = i
            while (--i >= 0) {
                if (data[i] == '\n') {
                    if (previousLine.replace(/\s*/,'').length > 0) {
                        if (previousLine.trim()[0] == '\'' || previousLine.trim()[0] == '<')
                            commentBlock = previousLine + '\n' + commentBlock;
                        else {
                            if ( commentBlock.length > 0 ){
                                // If we ate a blank line adjust the startpos to consume the character
                                startPos -= (1 - (maxBlankLines + 1));
                            }
                            break;
                        }
                    }else{
                        if ( ( commentBlock.length == 0 && maxBlankLines-- < -1 ) || commentBlock.length > 0 ){
                            break;
                        }
                    }
                    previousLine = '';
                } else {
                    previousLine = data[i] + previousLine;
                }
            }

            startPos -= (commentBlock.length + offset);

            if ( fnName.toLowerCase() === 'class_initialize' ){
                fnName = 'vb6_Class_Initialize';
                fnSignature = fnSignature.replace(/class_initialize/gi,fnName);
            }

            if ( fnName.toLowerCase() === 'class_terminate' ){
                fnName = 'vb6_Class_Terminate';
                fnSignature = fnSignature.replace(/class_terminate/gi,fnName);
            }

            //codeBlock =  + codeBlock;

            if (fnName !== undefined) {
                functionMap[fnName] =
                {   'name': fnName,
                    'type' : type,
                    'visibility' : visibility,
                    'global' : global,
                    'signature' : fnSignature,
                    'parameters' : parameters.trim().length > 0 ? parameters.split(',') : undefined,
                    'hits': 0
                };
            }

            var thisBlock = {
                'function' : functionMap[fnName],
                'comment' : commentBlock,
                'code' : sanitizer.clean( codeBlock.replace(/([^\n]+)/g, '\t$1') )
            };

            var result = variablesParser.parse( thisBlock.code, false, true );

            result.vars.forEach( function( thisVar ){
                if ( thisVar.name !== thisVar.var ){
                    thisBlock.code = thisBlock.code.substitute( thisVar.name, thisVar.var );
                }
            });

            thisBlock['vars'] = result.vars;

            functionBlocks.push( thisBlock );

            // Remove the comments and codeblock after the first \n
            var startStr = remainingData.substr(0,startPos - totalRemoved);
            var endStr = remainingData.substr(endPos - totalRemoved);

            totalRemoved += (endPos - startPos);

            remainingData = startStr + endStr;

            if ( remainingData.indexOf(fnSignature) != -1 ) {
                //console.log( 'processing error => ' + entry.in);
                console.log( "WARNING: " + fnSignature + ", still exists in the remaining data, could be a duplicate!. Source file => " + srcFile );
            }
        }

        return{
            'data' : remainingData.replace(/\n\n/gi, '\n').trim(),
            'map' : functionMap,
            'blocks' : functionBlocks
        }

    }
};

exports = module.exports = new FunctionsParser();