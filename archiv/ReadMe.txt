keywords: LCID, LocaleID, RFC1766, Ländercode, Sprachcode

Tipp-Upload: VB 5/6 0055: LCID in RFC1766-Zeichenfolge umwandeln

Der Vorschlag wurde erstellt am: 25.02.2008 15:25.

Die LCID oder auch LocaleID ist ein Länder- bzw Sprachencode der durch eine Integer-Zahl dargestellt wird. 
Diese Zahl setzt sich zusammen aus einem Anteil für die Sprache und einem Anteil für das jeweilige Land. 
Dabei ist jeder Kombination aus Land und Sprache eine andere Zahl zugeordnet. Für Deutschland ist das &H400, 
für Österreich &HC00. (Für andere Sprachen kann &H400 und &HC00 allerdings auch für andere Länder stehen) 
Für die Sprache deutsch, die beiden Ländern gemeinsam ist, ist das &H7 (die Zahl für die Sprache ist einmalig). 
In Kombination ergibt sich hieraus für deutsch in Deutschland &H407 und für deutsch in Österreich &HC07. 
Gebraucht werden kann so eine LocaleID im Zusammenhang mit länderspezifischen Einstellungen, wie Datumsformat, 
Währungsformat, etc. Da eine Zahl für Nichteingeweihte nicht sehr aussagekräftig ist, gibt es zu jeder LocaleID 
auch einen String-Schlüssel. Die beiden Funktionen wandeln eine LocaleID in einen aussagekräftigeren String um. 
Das gleiche Ergebnis könnte übrigens erreicht werden, wenn man in der Registry aus dem 
Schlüssel: "HKLM\SOFTWARE\Classes\MIME\Database\Rfc1766" alle (ca. 120) Einträge ausliest. 
(siehe auch Tipp: "Alle Sprachen auflisten")
Im .NET-Framework bietet die Klasse CultureInfo im Namespace System.Globalization die entsprechende Funktionalität. 
Über ihre Konstruktoren kann sie entweder mit einer LCID oder der RFC1766-Zeichenfolge initalisiert werden. 
Die Eigenschaft "Name" der Klasse liefert die RFC1766-Zeichenfolge, die Eigenschaft LCID die LCID als Integer zurück.
Die in diesem Tipp vorgestellten Funktionen machen dem Sinn nach dasselbe wie folgender Code in VB.NET:
MsgBox(New Globalization.CultureInfo("de-AT").LCID.ToString)
MsgBox(New Globalization.CultureInfo(&HC07).Name) 

VB 5/6-Tipp 0733: Alle Sprachen auflisten

Bezugnehmend auf Tippvorschlag "LCID in RFC1766-Zeichenfolge umwandeln", wird dieses Thema aufgegriffen um einen 
weiteren Tipp nachzulegen: 
Da man mit den Stringkürzeln, die die Funktion ConvertLCIDToRfc1766 liefert (z.B.: sv-FI), evtl. nicht gleich 
etwas anfangen kann, und man der FileVersioninfo-Klasse entnehmen kann, wie es besser geht, hier eine Funktion 
die einen LCID-Integer (bzw Long) in einen leserlichen Text verwandelt. Es werden in einer Unterroutine alle 
möglichen (und unmöglichen) LCID's durchlaufen (es werden auch LCIDs generiert, die keine Verwendung haben), 
mit der API-Funktion VerLanguageNameA in einen String konvertiert und in einer ListBox gesammelt.
