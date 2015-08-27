PropertiesParser = function(){

    var sanitizer = require('./code-sanitizer');
    var variablesParser = require('./variables-parser');

    this.parse = function(data){

        // Strip out any GET/LET statements
        var regEx = /(((?!'(?:\n))(?:\s*'.*?(?:\n))*)\n\s*((public|private)?\s*property\s+(get|let|set)\s+(\w+)\s*(?:\((\w*)\)){0,1})([\s\S]*?)(?:end\s+(?:property)))/gi;

        var remainingData = data;

        var classProperties = {};

        while (( match = regEx.exec(data) ) != null) {
            var codeBlock = match[0];

            remainingData = remainingData.replace(codeBlock, function(){
                return '';
            });

            var visibility = match[4];

            if (visibility === undefined)
                visibility = "Public";

            var property =
            {
                'comments' : ( match[2] != undefined ? match[2].replace(/\n\n/gi,'\n').replace(/^'/gm, '').split('\n') : undefined ),
                'visibility' : visibility,
                'type' : match[5].toLowerCase(), // get or let
                'name' : match[6],
                'code' : sanitizer.clean(match[8]),
                'function' : { 'name' : match[6] }
            };

            var result = variablesParser.parse( property.code );

            property['vars'] = result.vars;

            if ( classProperties[property.name]  === undefined )
                classProperties[property.name] = {
                    'function' : { 'name' : property.name },
                    'var' : property.name,
                    'name' : property.name
                };

            if ( property.type === 'set' )
                property.type == 'let';

            // Lets try and find a variable which is assigned so we can match it up with a DIM in the class
            if ( property.type === 'get' ){
                var regEx2 = new RegExp('\\s*\\b' + property.name + '\\b\\s*=\\s*(?:\\w*\\()?\\s*(\\w+)\\s*\\)?\\s*', 'gi');

                while (( match = regEx2.exec(property.code) ) != null) {
                    var theVar = match[1].trim();

                    if ( classProperties[property.name]['_Variable'] === undefined ){
                        remainingData = remainingData.replace( new RegExp('\\s*((private|public)(?:\\s*dim)?\\s+\\b' + theVar + '\\b)[\\t\\f]*(\'.*)?', 'gi' ), function( m, p1, p2, p3 ){
                            classProperties[property.name]['_Variable'] =
                            {
                                'name' : theVar,
                                'visibility' : p2,
                                'comment' : p3 !== undefined ? p3.replace(/^'/gm, '').split('\n') : undefined
                            };
                            return m;
                        });
                    }
                }
            }
            if ( property.type === 'let' ){
                var setParam = match[7];

                property['setParam'] = setParam;

                var regEx2 = new RegExp('\\s*(\\w+)\\b\\s*=\\s*(?:\\w*\\()?\\s*' + setParam + '\\s*\\)?\\s*', 'gi');

                while (( match = regEx2.exec(property.code) ) != null) {
                    var theVar = match[1].trim();

                    if ( classProperties[property.name]['_Variable'] === undefined ){
                        remainingData = remainingData.replace( new RegExp('\\s*((private|public)(?:\\s*dim)?\\s+\\b' + theVar + '\\b)[\\t\\f]*(\'.*)?', 'gi' ), function( m, p1, p2, p3 ){
                            classProperties[property.name]['_Variable'] =
                            {
                                'name' : theVar,
                                'visibility' : p2,
                                'comment' : p3 !== undefined ? p3.replace(/^'/gm, '').split('\n') : undefined
                            };
                            return m;
                        });
                    }
                }
            }

            classProperties[property.name][property.type] = property;

            //codeBlock = codeBlock.trim();

            //totalRemoved += codeBlock.length();
/*
            remainingData = remainingData.substr( match.index.toExponential()
            remainingData = remainingData.replace(codeBlock, function(m){
                return '';
            });
     */
        }

        classProperties.forEach = function (f){
            for( var name in this ) {
                if ( typeof this[name] === 'function' )
                    continue;
                f( this[name], name );
            }
        };

        return {
            'properties' : classProperties,
            'data' : remainingData.replace(/\n\n/gi, '\n').trim()
        }
    }

}

exports = module.exports = new PropertiesParser();