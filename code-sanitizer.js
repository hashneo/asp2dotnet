CodeSanitizer = function(){

    this.clean = function(code) {

        code = code//.replace(/Response\.Write\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Response.Write( $1 )$2')
            .replace(/Response\.Redirect\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Response.Redirect( $1 )$2')
            .replace(/Server\.Transfer\s+(.*(\s+&\s+vbCrLf)?)(.*)/gi,'Server.Transfer( $1 )$2')
            .replace(/if\s+err\s+then*/gi,'if Err.Number <> 0 Then')
            .replace(/(\s+)err(\s+)/gi,'$1Err.Number$2')
            .replace(/(\s+)set(\s+)/gi,'$1$2')
            .replace(/(\s+)wend(\s+)/gi,'$1End While$2')
            .replace(/\s{0}_$/gmi,' _')
            //.replace(/^\s*\w*\s*sub\s+class_initialize/gmi,'Sub vb6_Class_Initialize')
            //.replace(/^\s*\w*\s*sub\s+class_terminate/gmi,'Sub vb6_Class_Terminate')
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
};

exports = module.exports = new CodeSanitizer();