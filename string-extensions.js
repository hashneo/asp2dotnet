
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

String.prototype.getCommentsBlocks = function( pos ){

    var data = this;

    var i = pos;

    var offset = 0;
    // If we're on a \n shift to the first char
    while ( data[i + offset] == '\n' && i + offset < data.length ){
        offset++;
    }

    i += offset;

    var commentBlock = '';
    var previousLine = '';
    var maxBlankLines = 1;
    var startPos = i;
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

    if ( offset > 0 && commentBlock.length > offset )
        commentBlock = commentBlock.substr( 0, commentBlock.length - offset );

    return commentBlock;
}

String.prototype.remove = function(_what){

    if ( _what === undefined || _what === null || _what.length == 0 )
        return this;

    var removed = false;

    var data = this.replace(_what, function(){
        removed = true;
        return '';
    });

    if ( !removed ){
        throw "failed to remove string from data."
    }
    return data;
}

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

String.prototype.substitute = function( _what, _with, onMatch){

    var regEx;

    if ( typeof _what  === 'string' ){
        _what = {
            'name' : _what,
            'hits' : 0,
            'parameters' : []
        }
    }

    regEx = new RegExp('\\b((?:\\w+\\.)?' + _what.name.trim() + ')\\b.*?\n', 'gi');

    var code = this;

    code = code.replace( regEx, function(m, p1, offset){

        if ( onMatch !== undefined ){
            if ( !onMatch(p1) ){
                return m;
            }
        }

        // Check for being in comments or Strings
        var  x = offset;
        var p = 0;
        var c = 0;
        while(x >= 0 && code[x] != '\n'){
            if ( code[x] === '"' )
                p++;
            if ( ( p % 2 == 0 ) && code[x] === '\'' )
                c++;
            x--;
        }

        if ( p > 0 && p % 2 != 0 )
            return m;

        if ( c > 0 && c % 2 != 0 )
            return m;

        _what.hits += 1;

        if ( _what.parameters !== undefined && _what.parameters.length > 0 ){
            if (  m.match( new RegExp( p1 + '\\s*\\(') ) == null ){
                m = m.replace( new RegExp('\\b' + p1 + '\\b\\s*(?!\\s*\\=\\s*.*)(.*)', 'i'), ' ' + _with + '( $1 ) ' );
                return m;
            }
        }
        m = m.replace( new RegExp('\\b' + _what.name.trim() + '\\b', 'gi'), ' ' + _with + ' ' );
        return m;

    });

    return code;
};
