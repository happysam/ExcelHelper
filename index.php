<?php
header( "Content-type: text/html; charset=utf-8" );

for ($i = ord( 'K' ); $i <= ord( 'Z' ); $i++) 
{
	echo '<br>Ascii值:' . $i . ',字符:' . chr( $i );
}

echo '<br>K = ' . ord( 'K' );
echo '<br>Z = ' . ord( 'Z' );

$j = 0;
for ($i = 'K'; $i <= 'Z'; $i++)
{
	if ( 0 == $j % 2 ) 
	{
		echo '<br>Ascii值:' . $i . ',字符:' . ord( $i );
	}
	
	if ( 'BQ' === $i ) 
	{
		break;
	}
	
	$j ++;
}

