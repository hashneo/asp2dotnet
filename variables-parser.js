
VariablesParser = function() {

    this.parse = function(data, verbose) {

        function parseConst(code) {
            var comments = '';
            var match;



            if ( code.indexOf('mRenderTemplate') !== -1)
                var debug = 1;

            if (match = /((?:\s*'.*?\n)+)(?=\bconst\b\s+)/gi.exec(code)) {
                comments = match[1];
                code = code.replace(match[1], '');
            }

            // Sanity check to make sure we don't have any commented out variables
            if (code.search(/^\s*'.*?\bdim\b.*$/gmi) != -1)
                return null;

            if (match = /(private|public)?\s*const\s+(\w+)(?:\s*as\s*(\w+))?\s*=\s*(".*"|[\S\w]+)*\s*('.*)*/gi.exec(code)) {

                var visibility = match[1];

                if (visibility === undefined)
                    visibility = "Public";

                var result = {
                    'visibility': visibility,
                    'name': match[2],
                    'var': match[2],
                    'type': match[3],
                    'value': match[4],
                    'comment': ( match[5] != undefined ? match[5] : comments ).replace(/^'/gm, '').split('\n'),
                    'hits': 0
                };

                // Prevent short consts like 'rs' making into the wild
                if (result.name.length < 4) {
                    if (verbose)
                        console.log('WARNING: Const => ' + result.name + ' is less than 4 chars, marking it Private!');
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
            var match;


            if ( code.indexOf('g_xxx_pageobject_xxx') !== -1)
                var debug = 1;

            var start = code.search(/^\s*((public|private)\s+)?\b(?:dim)?\b/mi);

            comments = code.substr(0, start);
            code = code.substr(start);

            // Sanity check to make sure we don't have any commented out variables
            if (code.search(/^\s*'.*?\b(?:dim)?\b.*$/gmi) != -1)
                return null;

            if (comments.trim().length > 0) {
                comments = comments.replace(/^'/gm, '').split('\n');
                // If there are 2 previous blank lines before the comments, assume its junk and scrap the comment
                if (comments.length > 2 && comments[comments.length - 1].trim().length == 0 && comments[comments.length - 2].trim().length == 0)
                    comments = [];
            } else
                comments = [];

            if (match = /('.*)/gi.exec(code)) {
                comments = match[1].replace(/^'/gm, '').split('\n');
                code = code.replace(match[1], '');
            }

            code = code.replace(/[\n]/g, '');

            var results = [];

            var visibility;

            if (match = /(\s*(?:(public|private)?\s*(?:dim)?)\s*)/gi.exec(code)) {
                code = code.replace(match[1], '');
                visibility = match[2];
            }

            if (visibility === undefined)
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
                if (result.name.length < 4) {
                    /*
                    if (isNumeric(entry.name)) {
                        entry.name = '_' + entry.name;
                        if (verbose)
                            console.log('BIG WARNING: Variable => ' + result.name + ' has a numeric name!. I\'ll rename with and underscore to prevent mass updates');
                    }
                    */
                    if (verbose)
                        console.log('WARNING: Variable => ' + result.name + ' is less than 4 chars, marking it Private!');
                    result.visibility = 'Private';
                }


                results.push(result);
            }

            if (results.length == 0) {
                throw "unable to parse dim => " + code;
            }

            return results;
        }


        // Strip out all DIM and CONST declarations
        //var regEx = /^(((?:'.*?\n){0,}?)(?:private|public)?\s*\b(const+|dim?)\b\s+(?:[\s\S]+?))\s*$/gmi;
        var regEx = /^(((?:'.*?\n){0,}?)\s*(?:(?:private|public)?\s*\b(const|dim)|((?:private|public)+))\b\s+(?:[\s\S]+?))\s*$/gmi;

        var remainingData = data;

        var globalBlocks = '';
        var constDecls = [];
        var variableDecls = [];
        var match;

        while (( match = regEx.exec(data) ) != null) {
            var codeBlock = match[0];

            var endPos = match[0].length;
            if (codeBlock.replace('', '').trim().endsWith(',')) {
                codeBlock += '\n';
                var offset = match.index + match[0].length + 1;
                var pos;
                while ((pos = data.indexOf('\n', offset)) != -1) {
                    var nextLine = data.substr(offset, pos - offset);
                    offset += nextLine.length + 1;
                    endPos = offset;

                    codeBlock += nextLine + '\n';
                    if (!nextLine.replace('', '').trim().endsWith(','))
                        break;
                }
            }

            remainingData = remainingData.replace(codeBlock, "");

            switch (match[3] !== undefined ? match[3].toLowerCase() : '' ) {
                case 'dim':
                    var variableDecl = parseDim(codeBlock);
                    if (variableDecl != null)
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
                    var variableDecl = parseDim(codeBlock);
                    if (variableDecl != null)
                        variableDecls.push(variableDecl);
                    else
                        globalBlocks += '\t' + codeBlock;
                    break;
                    //globalBlocks += '\t' + codeBlock;
            }
        }

        variableDecls.forEach = function (f){
            for (var i = 0 ; i < this.length ; i++ ){
                for( var x = 0 ; x < this[i].length ; x++ ){
                    f( this[i][x], i, x );
                }
            }
        };

        constDecls.forEach = function (f){
            for (var i = 0 ; i < this.length ; i++ ){
                f( this[i], i );
            }
        };

        variableDecls.contains = function (_what){
            if ( this.length == 0 )
                return false;

            if ( _what === undefined )
                return false;

            _what = _what.toLowerCase();

            for (var i = 0 ; i < this.length ; i++ ){
                for( var x = 0 ; x < this[i].length ; x++ ){
                    if ( this[i][x].name.toLowerCase() === _what )
                        return true;
                }
            }

            return false;
        };

        return {
            'data': remainingData.replace(/\n\n/gi, '\n').trim(),
            'consts': constDecls,
            'vars': variableDecls,
            'unmatched': globalBlocks
        }
    }
}

exports = module.exports = new VariablesParser();