grammar TestExcel;


/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :
	LPAREN expression RPAREN #ParenthesizedExpr
	|expression EXPONENT expression #ExponentialExpr
    | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
    | expression operatorToken=(MODMATH | MOD | DIV) expression #ModDivExpr
	| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr


	| operatorToken=(MODMATH | MOD | DIV) LPAREN expression ',' expression RPAREN #ModDivExpr
	| MMAX LPAREN paramlist=arglist RPAREN #MmaxExpr
	| MMIN LPAREN paramlist=arglist RPAREN #MminExpr
	| MMAXSIMPLE LPAREN paramlist=arglist RPAREN #MaximumExpr
	| MMINSIMPLE LPAREN paramlist=arglist RPAREN #MinimumExpr
	| POW LPAREN paramlist=arglist RPAREN #PowExpr

	| NUMBER #NumberExpr
	| IDENTIFIER #IdentifierExpr
	; 

	arglist: expression (',' expression)+;
	paramlist: expression (',' expression)+;

/*
 * Lexer Rules
 */

NUMBER : INT ('.' INT)?; 
IDENTIFIER : [a-zA-Z]+[1-9][0-9]*;

INT : ('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
MOD : 'mod';
MODMATH : '%';
DIV : 'div';
MMAX : 'mmax';
MMIN : 'mmin';
MMAXSIMPLE : 'Maximum';
MMINSIMPLE : 'Minimum';
POW: 'pow';

WS : [ \t\r\n] -> channel(HIDDEN);
