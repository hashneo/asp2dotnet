CodeSanitizer = function(){

    this.clean = function(code) {

        code = code
            //.replace(/Response\.Redirect\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Response.Redirect( $1 )$2')
            //.replace(/Server\.Transfer\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Server.Transfer( $1 )$2')
            .replace(/if\s+err\s+then*/gi,'if Err.Number <> 0 Then')
            .replace(/(\s+)err(\s+)/gi,'$1Err.Number$2')
            .replace(/(\s+)set(\s+)/gi,'$1$2')
            .replace(/(\s+)wend(\s+)/gi,'$1End While$2')
            .replace(/(\s+)lenb(\s+)/gi,'$1len$2')
            .replace(/\s{0}_$/gmi,' _')
            .replace(/(\s+|\b)Server\.CreateObject\b\s*\((.*?)\)/gi,'$1CreateObject( $2 )')
            //.replace(/(\s+|\b)null\b(\s*)/gi,'$1VB6.Null$2')
            .replace(/(\s+|\b)Empty\b(\s*)/gi,'$1Nothing$2')
            //.replace(/(\s+|\b)Timer\b\s*\(\s*\)?(\s*)/gi,'$1New Timer()$2')
            .replace(/(\s+|\b)Date\b\s*\(\s*\)?(\s*)/gi,'$1New Date()$2')
            .replace(/Request\.QueryString\b(\s*\(.*?\))?(.*)/gi,'Request.QueryString$1.ToString()$2')
            .replace(/Request\.Cookies\b(\s*\(.*?\))?(.*)/gi,'Request.Cookies$1.ToString()$2')
            .replace(/\berr\.raise\b(\s+\w+.*)/gi,'Err.Raise($1)')

            .replace(/\bSetLocale\b\s*\(?(.*)\)?/gi,'Page.Culture = $1')
            .replace(/\bGetLocale\b\s*(?:\(\s*\))?/gi,'Page.Culture')


        code = code
            .replace(/Response\.Cookies\b\s*(\(.*?\))(\s*\(.*?\))?\s*=(.*)/gi, function(m,p1,p2,p3){
                if (p2 === undefined )
                    return 'Response.Cookies' + p1 + '.Value = ' + p3;
                return 'Response.Cookies' + p1 + '.Item' + p2 + ' = ' + p3;
        });

        code = code.replace(/Response\.CodePage\s*=\s*(\d+)/gi, 'Response.ContentEncoding = Encoding.GetEncoding($1)');

        var wordList = ['Open', 'Close', 'Read', 'Write', 'Add', 'AddNew', 'Update', 'Execute', 'Set', 'appendChild', 'setProperty', 'setNamedItem', 'appendChild', 'SaveToFile', 'LogEvent', 'WriteFile', 'AddHeader', 'Redirect', 'Transfer', 'Raise', 'Erase', 'Mark', 'MarkIn', 'MarkOut'];

        var regEx = new RegExp( '^(?!.*\\b\\w+\\.\\b(?:' + wordList.join('|' )+ ')\\s*(?:\\(.*\\)|=.*)$).*\\b(\\w+)\\.(' + wordList.join('|' )+ ')\\b\\s*?(.*)(\'.*)?$', 'gmi');

        code = code.replace(regEx, function(m, p1, p2, p3, p4 ){
            // Check for being in comments or Strings
            var x = m.indexOf(p1 + '.' + p2);
            var p = 0;
            var c = 0;
            while(x >= 0 && m[x] != '\n'){
                if ( m[x] === '"' )
                    p++;
                if ( ( p % 2 == 0 ) && m[x] === '\'' )
                    c++;
                x--;
            }

            if ( p > 0 && p % 2 != 0 )
                return m;

            if ( c > 0 && c % 2 != 0 )
                return m;

            if ( p3.indexOf('\'') != -1 ){
                x = p3.length - 1;
                while( x > 0 ){
                    if ( p3[x] == '\'' ){
                        x--;
                        return p1 + '.' + p2 + '(' + p3.substr(0,x) + ')' + p3.substr(x) + (p4 !== undefined ? p4 : '');
                    }
                    x--;
                }
                return p1 + '.' + p2 + '(' + p3 + ')' + (p4 !== undefined ? p4 : '');
            }
            else
                return p1 + '.' + p2 + '(' + p3 + ')' + (p4 !== undefined ? p4 : '');
        });

        code = code.replace(/\bNew\s+Date\(\)\s+(\+|\-)\s+(\d+)/gmi, 'New Date().AddDays($1$2)');

        //code = code.replace(/(\s+|\b)(CDate|Char)\b(\s*)/gi,'$1_$2$3')

        code = code.replace(/\bstring\b\s*\((.*?),(.*?)\)/gmi,'New String( $2, $1 )')

        // Nice replacements
        code = code.replace(/\.append\s+(.+?)(?='|$)/gmi,'.Append($1)');

        //.replace(/Append\s+(.*)/gi, 'Append( $1 ) ')

        //code = code.functionReplace( 'isobject', '($1 Is Nothing)' );

        code = code.functionReplace( 'round', 'Math.Round' );
        code = code.functionReplace( 'abs', 'Math.Abs' );
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
};

exports = module.exports = new CodeSanitizer();