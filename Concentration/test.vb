Option Strict Off
Option Explicit On
Module Cards32DLLAccess
	
	'Delarations and utilities for using Cards32.dll
	
	'Actions for CdtDraw/Ext
	' use in the nDraw field
	Public Const C_FACES As Short = 0
	Public Const C_BACKS As Short = 1
	Public Const C_INVERT As Short = 2
	
	'Card Numbers
	' use in the nCard field
	'from 0 to 51    [Ace (club,diamond,heart,spade), Deuce, ... , King]
	'Card Backs
	' use in the nCard field
	' CAUTION: when nCard > 53 then nDraw must be = 1 (C_BACKS)
	Public Const crosshatch As Short = 53
	Public Const weave As Short = 54
	Public Const tweed As Short = 55
	Public Const robot As Short = 56
	Public Const rose As Short = 57
	Public Const vineofthenight As Short = 58
	Public Const vineoftheday As Short = 59
	Public Const fourfish As Short = 60
	Public Const threefish As Short = 61
	'UPGRADE_NOTE: shell was upgraded to shell_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Public Const shell_Renamed As Short = 62
	Public Const castle As Short = 63
	Public Const island As Short = 64
	Public Const handfullofcards As Short = 65
	Public Const UNUSED As Short = 66
	Public Const THE_X As Short = 67
	Public Const THE_O As Short = 68
	
	'Initialization
	' call before anything else. Returns the default
	' width and height for the cards, in pixels.
	Declare Function cdtInit Lib "cards32.dll" (ByRef nWidth As Integer, ByRef nHeight As Integer) As Integer
	'Original:Declare Function cdtInit Lib "Cards32.dll" (nwidth As Integer, nheight As Integer) As Integer
	
	'CdtDraw used to draw a card with the default size
	'at a specified location in a form, picture box or whatever.
	'It can draw any of the 52 faces and 13 different Back designs,
	'as well as pile markers such as the X and O. Cards can also
	'be drawn in the negative image, eg to show selection.
	'xOrg = x origin in pixels
	'yOrg = y origin in pixels
	'nCard = one of the Card Back constants or a card number 0 to 51
	'nDraw = one of the Action constants 'nColor = The highlight color
	Declare Function cdtDraw Lib "cards32.dll" (ByVal hdc As Integer, ByVal xOrg As Integer, ByVal yOrg As Integer, ByVal nCard As Integer, ByVal nDraw As Integer, ByVal nColor As Integer) As Integer
	'Original: Declare Function cdtDraw Lib "Cards32.dll" (ByVal hDC As Integer, ByVal xOrg As Integer, ByVal yOrg As Integer, ByVal nCard As Integer, ByVal nDraw As Integer, ByVal nColor&) As Integer
	
	'CdtDrawExt used to draw a card in any size
	'Much the same as CdtDraw, but you can specify the height & width
	'of the card, as well as location.
	'nWidth = Width of card in pixels
	'nHeight = Height of card in pixels.
	Declare Function cdtDrawExt Lib "cards32.dll" (ByVal hdc As Integer, ByVal xOrg As Integer, ByVal yOrg As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal nCard As Integer, ByVal nDraw As Integer, ByVal nColor As Integer) As Integer
	'Original: Declare Function cdtDrawExt Lib "Cards32.dll" (ByVal hDC As Integer, ByVal xOrg As Integer, ByVal yOrg As Integer, ByVal nwidth As Integer, ByVal nheight As Integer, ByVal nCard As Integer, ByVal nDraw As Integer, ByVal nColor&) As Integer
	
	' cdtAnimate
	' Draws the animation on a card.  Four cards support animation:
	' cd         #frames    description
	' robot         4          robot meters
	' castle        2          bats flapping
	' island        4          sun sticks tongue out
	' cardhand      4          cards running up and down sleave
	' Card must be drawn at its default size
	' Call cdtAnimate every 250 ms for proper animation speed.
	' Arguments:
	'      HDC hdc
	'      int cd    robot, castle, island or cardhand
	'      int x     upper left corner of card
	'      int y
	'      int ispr  sprite to draw (0..1 for cdFaceDown10, 0..3 for others)
	' Returns:   TRUE if successful
	Declare Function cdtAnimate Lib "cards32.dll" (ByVal hdc As Integer, ByVal cd As Integer, ByVal xOrg As Integer, ByVal yOrg As Integer, ByVal ispr As Integer) As Integer
	
	'CdtTerm should be called when the program terminates.
	'Primarily it releases memory back to Windows.
	Declare Function cdtTerm Lib "cards32.dll" () As Integer
	'Original: Declare Function cdtTerm Lib "Cards32.dll" () As Integer
End Module