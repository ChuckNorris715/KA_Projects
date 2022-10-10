FasdUAS 1.101.10   ��   ��    k             l     ��  ��    h b##################################################################################################     � 	 	 � # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #   
  
 l     ��  ��    p j#	This AppleScript is for generating Mail Signature Templates in Microsoft Outlook for Mac														##     �   � # 	 T h i s   A p p l e S c r i p t   i s   f o r   g e n e r a t i n g   M a i l   S i g n a t u r e   T e m p l a t e s   i n   M i c r o s o f t   O u t l o o k   f o r   M a c 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # #      l     ��  ��    g a#	Prepared by Stuart Lamont in November 2015 to replace the Centenary Signatures															##     �   � # 	 P r e p a r e d   b y   S t u a r t   L a m o n t   i n   N o v e m b e r   2 0 1 5   t o   r e p l a c e   t h e   C e n t e n a r y   S i g n a t u r e s 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # #      l     ��  ��    * $#																																	##     �   H # 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # #      l     ��  ��    l f#	The script performs numerous Active Directory Lookups and will produce inconsistent															##     �   � # 	 T h e   s c r i p t   p e r f o r m s   n u m e r o u s   A c t i v e   D i r e c t o r y   L o o k u p s   a n d   w i l l   p r o d u c e   i n c o n s i s t e n t 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # #      l     ��   !��     _ Y#	results if the Active Directory binding is in any way compromised.																			##    ! � " " � # 	 r e s u l t s   i f   t h e   A c t i v e   D i r e c t o r y   b i n d i n g   i s   i n   a n y   w a y   c o m p r o m i s e d . 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # #   # $ # l     �� % &��   % n h#	If Surname, Title and Phone Number aren't on the Generated Template, unbind and re-bind													##    & � ' ' � # 	 I f   S u r n a m e ,   T i t l e   a n d   P h o n e   N u m b e r   a r e n ' t   o n   t h e   G e n e r a t e d   T e m p l a t e ,   u n b i n d   a n d   r e - b i n d 	 	 	 	 	 	 	 	 	 	 	 	 	 # # $  ( ) ( l     �� * +��   * [ U#	the computer with Active Directory and the re-run the script.																				##    + � , , � # 	 t h e   c o m p u t e r   w i t h   A c t i v e   D i r e c t o r y   a n d   t h e   r e - r u n   t h e   s c r i p t . 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # # )  - . - l     �� / 0��   / * $#																																	##    0 � 1 1 H # 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # # .  2 3 2 l     �� 4 5��   4 m g#	If the Script is run more than once, multiple Templates will be generated, so please															##    5 � 6 6 � # 	 I f   t h e   S c r i p t   i s   r u n   m o r e   t h a n   o n c e ,   m u l t i p l e   T e m p l a t e s   w i l l   b e   g e n e r a t e d ,   s o   p l e a s e 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # # 3  7 8 7 l     �� 9 :��   9 ` Z#	bear this in mind when selecting the default templates for the user.																		##    : � ; ; � # 	 b e a r   t h i s   i n   m i n d   w h e n   s e l e c t i n g   t h e   d e f a u l t   t e m p l a t e s   f o r   t h e   u s e r . 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 # # 8  < = < l     �� > ?��   > h b##################################################################################################    ? � @ @ � # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # =  A B A l     ��������  ��  ��   B  C D C l     �� E F��   E ( "set MyName to name of me as string    F � G G D s e t   M y N a m e   t o   n a m e   o f   m e   a s   s t r i n g D  H I H l     �� J K��   J  display dialog MyName    K � L L * d i s p l a y   d i a l o g   M y N a m e I  M N M l     �� O P��   O " instantiate global variables    P � Q Q 8 i n s t a n t i a t e   g l o b a l   v a r i a b l e s N  R S R l     �� T U��   T A ;global variables are used here for the subroutine to access    U � V V v g l o b a l   v a r i a b l e s   a r e   u s e d   h e r e   f o r   t h e   s u b r o u t i n e   t o   a c c e s s S  W X W p       Y Y ������ 0 longname longName��   X  Z [ Z p       \ \ ������ 0 username userName��   [  ] ^ ] p       _ _ ������ 0 
rawsurname  ��   ^  ` a ` p       b b ������ 0 	firstname  ��   a  c d c p       e e ������ 0 surname  ��   d  f g f p       h h ������ 0 	nametitle  ��   g  i j i p       k k ������ 	0 email  ��   j  l m l p       n n ������ 0 jobtitle jobTitle��   m  o p o p       q q ������ 0 phoneno phoneNo��   p  r s r p       t t ������ 0 directphone directPhone��   s  u v u p       w w ������ 0 address1  ��   v  x y x p       z z ������ 0 	descript1  ��   y  { | { p       } } ������ 0 	descript2  ��   |  ~  ~ p       � � ������ 0 fontcolour1 fontColour1��     � � � p       � � ������ 0 fontcolour2 fontColour2��   �  � � � p       � � ������ 0 location1name location1Name��   �  � � � p       � � ������ 0 location2name location2Name��   �  � � � p       � � ������ 0 descriptmain descriptMain��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � # Variables for Graphics Assets    � � � � : V a r i a b l e s   f o r   G r a p h i c s   A s s e t s �  � � � p       � � ������ 0 logolink logoLink��   �  � � � p       � � ������ 0 weburl webURL��   �  � � � p       � � ������ 0 
weburltext 
webURLText��   �  � � � p       � � ������ 0 twitterlink twitterLink��   �  � � � p       � � ������ "0 twitterlogolink twitterLogoLink��   �  � � � p       � � ������ 0 facebooklink facebookLink��   �  � � � p       � � ������ $0 facebooklogolink facebookLogoLink��   �  � � � p       � � ������ 0 linkedinlink linkedInLink��   �  � � � p       � � ������ $0 linkedinlogolink linkedInLogoLink��   �  � � � p       � � ������ 0 	instalink 	instaLink��   �  � � � p       � � ������ 0 instalogolink instaLogoLink��   �  � � � p       � � ������ &0 bottomborderimage bottomBorderImage��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   �  variable for HTML Block    � � � � . v a r i a b l e   f o r   H T M L   B l o c k �  � � � p       � � ������ 0 htmlcontent HTMLContent��   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � 8 2Collect User Data and place in Variable containers    � � � � d C o l l e c t   U s e r   D a t a   a n d   p l a c e   i n   V a r i a b l e   c o n t a i n e r s �  � � � l     ����� � O      � � � k     � �  � � � r     � � � 1    ��
�� 
siln � o      ���� 0 longname longName �  ��� � r     � � � 1    ��
�� 
sisn � o      ���� 0 username userName��   � l     ����� � e      � � I    ������
�� .sysosigtsirr   ��� null��  ��  ��  ��  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � ' !pull Surname from System LongName    � � � � B p u l l   S u r n a m e   f r o m   S y s t e m   L o n g N a m e �  � � � l     �� � ���   �  not used    � � � �  n o t   u s e d �  � � � l   , ����� � r    , � � � n    * � � � 7   *�� � �
�� 
ctxt � m    ����  � l   ) ����� � \    ) � � � l   ' ����� � I   '���� �
�� .sysooffslong    ��� null��   � �� � �
�� 
psof � m     ! � � � � �  , � �� ���
�� 
psin � o   " #���� 0 longname longName��  ��  ��   � m   ' (�� ��  ��   � o    �~�~ 0 longname longName � o      �}�} 0 
rawsurname  ��  ��   �  � � � l     �| � ��|   � * $pull first name from System LongName    � � � � H p u l l   f i r s t   n a m e   f r o m   S y s t e m   L o n g N a m e �  � � � l  - 4 �{�z  r   - 4 I  - 2�y�x
�y .sysoexecTEXT���     TEXT l  - .�w�v m   - . � � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   g i v e n N a m e   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�w  �v  �x   o      �u�u 0 	firstname  �{  �z   �  l     �t	
�t  	 0 *Pull Surname from AD Extension Attribute 1   
 � T P u l l   S u r n a m e   f r o m   A D   E x t e n s i o n   A t t r i b u t e   1  l  5 >�s�r r   5 > I  5 :�q�p
�q .sysoexecTEXT���     TEXT l  5 6�o�n m   5 6 � � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   s n   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�o  �n  �p   o      �m�m 0 surname  �s  �r    l     �l�l   B <pull pull "Title" (Dr/Rev/etc) from AD Extension Attribute 3    � x p u l l   p u l l   " T i t l e "   ( D r / R e v / e t c )   f r o m   A D   E x t e n s i o n   A t t r i b u t e   3  l  ? J�k�j r   ? J I  ? F�i�h
�i .sysoexecTEXT���     TEXT l  ? B �g�f  m   ? B!! �"" � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   d e s c r i p t i o n   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�g  �f  �h   o      �e�e 0 	nametitle  �k  �j   #$# l     �d%&�d  % 6 0pull email address from AD Extension Attribute 2   & �'' ` p u l l   e m a i l   a d d r e s s   f r o m   A D   E x t e n s i o n   A t t r i b u t e   2$ ()( l  K V*�c�b* r   K V+,+ I  K R�a-�`
�a .sysoexecTEXT���     TEXT- l  K N.�_�^. m   K N// �00 � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   m a i l   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�_  �^  �`  , o      �]�] 	0 email  �c  �b  ) 121 l     �\34�\  3 / )pull job title from AD Attribute JobTitle   4 �55 R p u l l   j o b   t i t l e   f r o m   A D   A t t r i b u t e   J o b T i t l e2 676 l  W b8�[�Z8 r   W b9:9 I  W ^�Y;�X
�Y .sysoexecTEXT���     TEXT; l  W Z<�W�V< m   W Z== �>> � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   t i t l e   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�W  �V  �X  : o      �U�U 0 jobtitle jobTitle�[  �Z  7 ?@? l     �TAB�T  A ? 9pull telephone Extension number from AD Attribute ipPhone   B �CC r p u l l   t e l e p h o n e   E x t e n s i o n   n u m b e r   f r o m   A D   A t t r i b u t e   i p P h o n e@ DED l  c nF�S�RF r   c nGHG I  c j�QI�P
�Q .sysoexecTEXT���     TEXTI l  c fJ�O�NJ m   c fKK �LL � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   t e l e p h o n e N u m b e r   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�O  �N  �P  H o      �M�M 0 phoneno phoneNo�S  �R  E MNM l     �LOP�L  O 5 /pull direct telephone from AD Attribute ipPhone   P �QQ ^ p u l l   d i r e c t   t e l e p h o n e   f r o m   A D   A t t r i b u t e   i p P h o n eN RSR l  o zT�K�JT r   o zUVU I  o v�IW�H
�I .sysoexecTEXT���     TEXTW l  o rX�G�FX m   o rYY �ZZ � " / A p p l i c a t i o n s / E n t e r p r i s e   C o n n e c t . a p p / C o n t e n t s / S h a r e d S u p p o r t / e c c l "   - a   i p P h o n e   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '�G  �F  �H  V o      �E�E 0 directphone directPhone�K  �J  S [\[ l     �D�C�B�D  �C  �B  \ ]^] l     �A_`�A  _ : 4####################################################   ` �aa h # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #^ bcb l     �@de�@  d  Setup Addresses   e �ff  S e t u p   A d d r e s s e sc ghg l  { �i�?�>i r   { �jkj m   { ~ll �mm L 7 5 6   H a d d o n   A v e .   C o l l i n g s w o o d ,   N J   0 8 1 0 8k o      �=�= 0 address1  �?  �>  h non l     �<�;�:�<  �;  �:  o pqp l     �9�8�7�9  �8  �7  q rsr l     �6tu�6  t : 4####################################################   u �vv h # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #s wxw l     �5yz�5  y  Setup Font Colours   z �{{ $ S e t u p   F o n t   C o l o u r sx |}| l  � �~�4�3~ r   � �� b   � ���� b   � ���� m   � ��� ���� < t r > 
                                   < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 1 1 p x ; l i n e - h e i g h t : 1 5 p x ; c o l o r : # 2 3 1 f 2 0 ;   f o n t - w e i g h t : 6 0 0 ; p a d d i n g - t o p : 6 p x ; " >  � o   � ��2�2 0 	nametitle  � m   � ��� ��� 4 < / t d > 
                               < / t r >� o      �1�1 0 	descript1  �4  �3  } ��� l  � ���0�/� r   � ���� m   � ��� ���  � o      �.�. 0 	descript2  �0  �/  � ��� l     �-�,�+�-  �,  �+  � ��� l     �*���*  � : 4####################################################   � ��� h # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #� ��� l     �)���)  �  Setup Location Names   � ��� ( S e t u p   L o c a t i o n   N a m e s� ��� l     �(���(  � $ set location1Name to "Ivanhoe"   � ��� < s e t   l o c a t i o n 1 N a m e   t o   " I v a n h o e "� ��� l     �'���'  � # set location2Name to "Plenty"   � ��� : s e t   l o c a t i o n 2 N a m e   t o   " P l e n t y "� ��� l     �&�%�$�&  �%  �$  � ��� l     �#���#  � : 4####################################################   � ��� h # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #� ��� l     �"���"  �  setup graphical Assets   � ��� , s e t u p   g r a p h i c a l   A s s e t s� ��� l  � ���!� � r   � ���� m   � ��� ��� t h t t p : / / m e d i a . i g s . v i c . e d u . a u / s i g n a t u r e s / I v a n h o e L i n e B i g 2 . p n g� o      �� 0 logolink logoLink�!  �   � ��� l  � ����� r   � ���� m   � ��� ��� 2 h t t p : / / w w w . i v a n h o e . c o m . a u� o      �� 0 weburl webURL�  �  � ��� l  � ����� r   � ���� m   � ��� ��� $ w w w . i v a n h o e . c o m . a u� o      �� 0 
weburltext 
webURLText�  �  � ��� l  � ����� r   � ���� m   � ��� ��� B h t t p : / / t w i t t e r . c o m / i v a n h o e g r a m m a r� o      �� 0 twitterlink twitterLink�  �  � ��� l  � ����� r   � ���� m   � ��� ��� j h t t p : / / m e d i a . i g s . v i c . e d u . a u / s i g n a t u r e s / t w i t t e r s m l . p n g� o      �� "0 twitterlogolink twitterLogoLink�  �  � ��� l  � ����� r   � ���� m   � ��� ��� X h t t p : / / w w w . f a c e b o o k . c o m / I v a n h o e G r a m m a r S c h o o l� o      �� 0 facebooklink facebookLink�  �  � ��� l  � ����� r   � ���� m   � ��� ��� l h t t p : / / m e d i a . i g s . v i c . e d u . a u / s i g n a t u r e s / f a c e b o o k s m l . p n g� o      �� $0 facebooklogolink facebookLogoLink�  �  � ��� l  � ����� r   � ���� m   � ��� ��� n h t t p s : / / w w w . l i n k e d i n . c o m / c o m p a n y / i v a n h o e - g r a m m a r - s c h o o l� o      �
�
 0 linkedinlink linkedInLink�  �  � ��� l  � ���	�� r   � ���� m   � ��� ��� l h t t p : / / m e d i a . i g s . v i c . e d u . a u / s i g n a t u r e s / l i n k e d i n s m l . p n g� o      �� $0 linkedinlogolink linkedInLogoLink�	  �  � ��� l  � ����� r   � ���� m   � ��� ��� ^ h t t p s : / / w w w . i n s t a g r a m . c o m / i v a n h o e g r a m m a r s c h o o l /� o      �� 0 	instalink 	instaLink�  �  � ��� l  � ����� r   � ���� m   � ��� ��� n h t t p : / / m e d i a . i g s . v i c . e d u . a u / s i g n a t u r e s / i n s t a g r a m s m l . p n g� o      �� 0 instalogolink instaLogoLink�  �  � � � l  � �� �� r   � � m   � � � ~ h t t p : / / m e d i a . i g s . v i c . e d u . a u / g e n e r a l / s i g n a t u r e s / b o t t o m b o r d e r . j p g o      ���� &0 bottomborderimage bottomBorderImage�   ��     l     ��������  ��  ��   	 l     ��
��  
  Error Checking    �  E r r o r   C h e c k i n g	  l     ����   g acheck for field data complete - If surname is Blank, quit, and prompt user to come to IT Services    � � c h e c k   f o r   f i e l d   d a t a   c o m p l e t e   -   I f   s u r n a m e   i s   B l a n k ,   q u i t ,   a n d   p r o m p t   u s e r   t o   c o m e   t o   I T   S e r v i c e s  l     ����    if surname is "" then    � * i f   s u r n a m e   i s   " "   t h e n  l     ����   � �	display dialog "This Action cannot be completed as your computer's Active Directory Binding is broken. Please bring your computer to IT Services to correct this issue." with icon stop buttons "Exit"    �� 	 d i s p l a y   d i a l o g   " T h i s   A c t i o n   c a n n o t   b e   c o m p l e t e d   a s   y o u r   c o m p u t e r ' s   A c t i v e   D i r e c t o r y   B i n d i n g   i s   b r o k e n .   P l e a s e   b r i n g   y o u r   c o m p u t e r   t o   I T   S e r v i c e s   t o   c o r r e c t   t h i s   i s s u e . "   w i t h   i c o n   s t o p   b u t t o n s   " E x i t "  l     ����    	return    �    	 r e t u r n !"! l     ��������  ��  ��  " #$# l     ��%&��  %  end if   & �''  e n d   i f$ ()( l     ��������  ��  ��  ) *+* l     ��������  ��  ��  + ,-, l     ��������  ��  ��  - ./. l     ��01��  0 � �##############################################################################################################################################################   1 �22< # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #/ 343 l     ��������  ��  ��  4 565 l     ��78��  7 f `Prompt user to select which Campus they are based at. Will determine which template is generated   8 �99 � P r o m p t   u s e r   t o   s e l e c t   w h i c h   C a m p u s   t h e y   a r e   b a s e d   a t .   W i l l   d e t e r m i n e   w h i c h   t e m p l a t e   i s   g e n e r a t e d6 :;: l     ��<=��  < � �set question to display dialog "Which Campus are you based at?" buttons {location1Name, location2Name} default button location1Name   = �>> s e t   q u e s t i o n   t o   d i s p l a y   d i a l o g   " W h i c h   C a m p u s   a r e   y o u   b a s e d   a t ? "   b u t t o n s   { l o c a t i o n 1 N a m e ,   l o c a t i o n 2 N a m e }   d e f a u l t   b u t t o n   l o c a t i o n 1 N a m e; ?@? l     ��AB��  A / )set campus to button returned of question   B �CC R s e t   c a m p u s   t o   b u t t o n   r e t u r n e d   o f   q u e s t i o n@ DED l     ��FG��  F  Ridgeway Template   G �HH " R i d g e w a y   T e m p l a t eE IJI l  �&K����K Z   �&LM��NL =  �OPO o   � ����� 0 	nametitle  P m   �QQ �RR  M k  SS TUT l ��������  ��  ��  U VWV r  XYX o  ���� 0 	descript2  Y o      ���� 0 descriptmain descriptMainW Z[Z I  ��������  0 setupsignature setupSignature��  ��  [ \��\ l ��������  ��  ��  ��  ��  N k  &]] ^_^ l ��������  ��  ��  _ `a` r  bcb o  ���� 0 	descript1  c o      ���� 0 descriptmain descriptMaina ded I  $��������  0 setupsignature setupSignature��  ��  e fgf l %%��������  ��  ��  g h��h l %%��������  ��  ��  ��  ��  ��  J iji l     ��������  ��  ��  j klk i     mnm I      ��������  0 setupsignature setupSignature��  ��  n k     aoo pqp O     _rsr I   ^����t
�� .corecrel****      � null��  t ��uv
�� 
koclu m   
 ��
�� 
cSigv ��w��
�� 
prdtw K    Xxx ��yz
�� 
pnamy m    {{ �|| , N E W _ D O M A I N _ W O _ D e s c r i p tz ��}��
�� 
ctnt} b    T~~ b    P��� b    L��� b    H��� b    D��� b    @��� b    <��� b    8��� b    4��� b    0��� b    ,��� b    (��� b    $��� b     ��� b    ��� b    ��� b    ��� b    ��� b    ��� b    ��� m    �� ���f < h t m l > 
 < b o d y   c l a s s = " q e _ b o d y "   s t y l e = " p a d d i n g : 0 ;   m a r g i n : 0   a u t o   ! i m p o r t a n t ;   d i s p l a y : b l o c k   ! i m p o r t a n t ;   m i n - w i d t h : 1 0 0 %   ! i m p o r t a n t ;   w i d t h : 1 0 0 %   ! i m p o r t a n t ;   b a c k g r o u n d : # f f f f f f ;   - w e b k i t - t e x t - s i z e - a d j u s t : n o n e " > 
 < t a b l e   w i d t h = " 1 0 0 % "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 "   b g c o l o r = " # f f f f f f "     c l a s s = " f u l l - w r a p " > 
     < t r > 
         < t d   a l i g n = " c e n t e r "   v a l i g n = " t o p " > < t a b l e   a l i g n = " l e f t "   s t y l e = " w i d t h : 3 2 0 p x ;   m a x - w i d t h : 3 2 0 p x ;   t a b l e - l a y o u t : f i x e d ; "   c l a s s = " q e _ w r a p p e r "     w i d t h = " 3 2 0 "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 " > 
                 < t r > 
                     < t d   v a l i g n = " t o p "   a l i g n = " c e n t e r "   s t y l e = " p a d d i n g : 2 0 p x   6 p x ; " > < t a b l e   w i d t h = " 1 0 0 % "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 "   a l i g n = " c e n t e r " > 
                             < t r > 
                                 < t d   v a l i g n = " m i d d l e "   a l i g n = " l e f t "   w i d t h = " 1 0 4 "   s t y l e = " w i d t h : 1 0 4 p x ; p a d d i n g - t o p : 4 p x ; " > < a   h r e f = " h t t p s : / / w w w . t h r i v e n . d e s i g n / "   t a r g e t = " _ b l a n k "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; " > < i m g   s r c = " h t t p : / / z e r o z o n e . c o m / q e i n b o x / s i g n a t u r e s / l o g o _ u p d a t e d . p n g "   w i d t h = " 1 0 4 "   a l t = " t h r i v e n   d e s i g n "   b o r d e r = " 0 "   s t y l e = " f o n t - f a m i l y : A r i a l ,   s a n s - s e r i f ;   f o n t - s i z e : 1 4 p x ;   l i n e - h e i g h t : 1 7 p x ; c o l o r : # 0 0 0 0 0 0 ; d i s p l a y : b l o c k ; m a x - w i d t h : 1 0 4 p x ; " / > < / a > < / t d > 
                                 < t d   v a l i g n = " m i d d l e "   a l i g n = " c e n t e r "   s t y l e = " p a d d i n g - l e f t : 1 5 p x ; " > < t a b l e   w i d t h = " 1 0 0 % "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 " > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 1 6 p x ; l i n e - h e i g h t : 2 0 p x ; c o l o r : # 2 3 1 f 2 0 ;   f o n t - w e i g h t : b o l d ; " >� o    ���� 0 	firstname  � m    �� ���  & n b s p ;� o    ���� 0 surname  � m    �� ��� H < / t d > 
                                         < / t r > 
 	 	 	 	� o    ���� 0 descriptmain descriptMain� m    �� ���� 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 1 0 p x ; l i n e - h e i g h t : 1 3 p x ; c o l o r : # 0 0 0 0 0 0 ;   p a d d i n g - t o p : 5 p x ; " >� o    ���� 0 jobtitle jobTitle� m     #�� ���B < / t d > 
                                         < / t r > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " c e n t e r "   s t y l e = " p a d d i n g : 5 p x   0 p x ; " > < t a b l e   w i d t h = " 1 0 0 % "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 "   a l i g n = " l e f t "   > 
                                                     < t r > 
                                                         < t d   h e i g h t = " 1 "   s t y l e = " h e i g h t : 1 p x ; f o n t - s i z e : 0 p x ; l i n e - h e i g h t : 0 p x ; "   b g c o l o r = " # 0 0 0 0 0 0 " > < / t d > 
                                                     < / t r > 
                                                 < / t a b l e > < / t d > 
                                         < / t r > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 9 p x ; l i n e - h e i g h t : 1 3 p x ; c o l o r : # 0 0 0 0 0 0 ; " > < a   h r e f = " m a i l t o :� o   $ '���� 	0 email  � m   ( +�� ��� \ "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; c o l o r : # 0 0 0 0 0 0 ; " >� o   , /���� 	0 email  � m   0 3�� ���: < / a > < / t d > 
                                         < / t r > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 9 p x ; l i n e - h e i g h t : 1 3 p x ; c o l o r : # 0 0 0 0 0 0 ; p a d d i n g - t o p : 5 p x ;   " > < s t r o n g > T :   < / s t r o n g > < a   h r e f = " t e l :� o   4 7���� 0 phoneno phoneNo� m   8 ;�� ��� \ "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; c o l o r : # 0 0 0 0 0 0 ; " >� o   < ?���� 0 phoneno phoneNo� m   @ C�� ��� d < / a > & n b s p ; | & n b s p ; < s t r o n g > D :   < / s t r o n g > < a   h r e f = " t e l :� o   D G���� 0 directphone directPhone� m   H K�� ��� \ "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; c o l o r : # 0 0 0 0 0 0 ; " >� o   L O���� 0 directphone directPhone m   P S�� ���� < / a > < / t d > 
                                         < / t r > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 9 p x ; l i n e - h e i g h t : 1 3 p x ; c o l o r : # 0 0 0 0 0 0 ;   p a d d i n g - t o p : 5 p x ; " > 7 5 6   H a d d o n   A v e .   C o l l i n g s w o o d ,   N J   0 8 1 0 8 < / t d > 
                                         < / t r > 
                                         < t r > 
                                             < t d   v a l i g n = " t o p "   a l i g n = " c e n t e r "   s t y l e = " p a d d i n g - t o p : 6 p x ; " > < t a b l e   w i d t h = " 1 0 0 % "   b o r d e r = " 0 "   c e l l s p a c i n g = " 0 "   c e l l p a d d i n g = " 0 "   a l i g n = " c e n t e r " > 
                                                     < t r > 
                                                         < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   w i d t h = " 2 6 "   s t y l e = " w i d t h : 2 6 p x ;   l i n e - h e i g h t : 0 p x ;   f o n t - s i z e : 0 p x ; " > < a   h r e f = " h t t p s : / / w w w . l i n k e d i n . c o m / c o m p a n y / t h r i v e n . d e s i g n "   t a r g e t = " _ b l a n k "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; " > < i m g   s r c = " h t t p : / / z e r o z o n e . c o m / q e i n b o x / s i g n a t u r e s / l i n k e d i n . p n g "   w i d t h = " 2 2 "     b o r d e r = " 0 "   s t y l e = " f o n t - f a m i l y : A r i a l ,   s a n s - s e r i f ;   f o n t - s i z e : 1 4 p x ;   l i n e - h e i g h t : 1 7 p x ; c o l o r : # 0 0 0 0 0 0 ; d i s p l a y : b l o c k ; m a x - w i d t h : 2 2 p x ; " / > < / a > < / t d > 
                                                         < t d   w i d t h = " 5 "   s t y l e = " w i d t h : 5 p x ; l i n e - h e i g h t : 0 p x ; f o n t - s i z e : 0 p x ; " > < / t d > 
                                                         < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   w i d t h = " 2 7 "   s t y l e = " w i d t h : 2 7 p x ;   l i n e - h e i g h t : 0 p x ;   f o n t - s i z e : 0 p x ; " > < a   h r e f = " h t t p s : / / w w w . i n s t a g r a m . c o m / t h r i v e n . d e s i g n / "   t a r g e t = " _ b l a n k "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; " > < i m g   s r c = " h t t p : / / z e r o z o n e . c o m / q e i n b o x / s i g n a t u r e s / i n s t a . p n g "   w i d t h = " 2 2 "     b o r d e r = " 0 "   s t y l e = " f o n t - f a m i l y : A r i a l ,   s a n s - s e r i f ;   f o n t - s i z e : 1 4 p x ;   l i n e - h e i g h t : 1 7 p x ; c o l o r : # 0 0 0 0 0 0 ; d i s p l a y : b l o c k ; m a x - w i d t h : 2 2 p x ; " / > < / a > < / t d > 
                                                         < t d   w i d t h = " 5 "   s t y l e = " w i d t h : 5 p x ; l i n e - h e i g h t : 0 p x ; f o n t - s i z e : 0 p x ; " > < / t d > 
                                                         < t d   v a l i g n = " t o p "   a l i g n = " l e f t "   w i d t h = " 2 7 "   s t y l e = " w i d t h : 2 7 p x ;   l i n e - h e i g h t : 0 p x ;   f o n t - s i z e : 0 p x ; " > < a   h r e f = " h t t p s : / / w w w . f a c e b o o k . c o m / t h r i v e n . d e s i g n / "   t a r g e t = " _ b l a n k "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; " > < i m g   s r c = " h t t p : / / z e r o z o n e . c o m / q e i n b o x / s i g n a t u r e s / f a c e b o o k . p n g "   w i d t h = " 2 2 "     b o r d e r = " 0 "   s t y l e = " f o n t - f a m i l y : A r i a l ,   s a n s - s e r i f ;   f o n t - s i z e : 1 4 p x ;   l i n e - h e i g h t : 1 7 p x ; c o l o r : # 0 0 0 0 0 0 ; d i s p l a y : b l o c k ; m a x - w i d t h : 2 2 p x ; " / > < / a > < / t d > 
                                                         < t d   w i d t h = " 1 0 "   s t y l e = " w i d t h : 1 0 p x ; " > < / t d > 
                                                         < t d   v a l i g n = " m i d d l e "   a l i g n = " l e f t "   c l a s s = " q e _ d e f a u l t l i n k "   s t y l e = " f o n t - f a m i l y :   ' M o n t s e r r a t ' ,   A r i a l ,   s a n s - s e r i f ; f o n t - s i z e : 1 0 p x ; l i n e - h e i g h t : 1 3 p x ; c o l o r : # 0 0 0 0 0 0 ; f o n t - w e i g h t : 6 0 0 ;   " > < a   h r e f = " h t t p s : / / w w w . t h r i v e n . d e s i g n / "   t a r g e t = " _ b l a n k "   s t y l e = " t e x t - d e c o r a t i o n : n o n e ; c o l o r : # 0 0 0 0 0 0 ; " > t h r i v e n . d e s i g n < / a > < / t d > 
                                                     < / t r > 
                                                 < / t a b l e > < / t d > 
                                         < / t r > 
                                     < / t a b l e > < / t d > 
                             < / t r > 
                         < / t a b l e > < / t d > 
                 < / t r > 
             < / t a b l e > < / t d > 
     < / t r > 
 < / t a b l e > 
 < / b o d y > 
 < / h t m l >��  ��  s 5     �����
�� 
capp� m    �� ��� * c o m . m i c r o s o f t . O u t l o o k
�� kfrmID  q ���� l  ` `��������  ��  ��  ��  l ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � ? 9#########################################################   � ��� r # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #� ��� l     ������  � f `This subRoutine currrently will not function without "Accessibility Access" enabled for the app.   � ��� � T h i s   s u b R o u t i n e   c u r r r e n t l y   w i l l   n o t   f u n c t i o n   w i t h o u t   " A c c e s s i b i l i t y   A c c e s s "   e n a b l e d   f o r   t h e   a p p .� ���� i    ��� I      ������� $0 updatedefaultsig updateDefaultSig��  � ����
�� 
to  � o      ���� 0 mysignature mySignature� �����
�� 
for � o      ���� 0 accountname accountName��  � k    ��� ��� O     
��� I   	������
�� .miscactvnull��� ��� null��  ��  � m     ���                                                                                  OPIM  alis    N  Macintosh HD               �=�wBD ����Microsoft Outlook.app                                          ����ޛ�0        ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  � ��� l   ��������  ��  ��  � ��� O    8��� O    7��� O    6��� O    5��� O   $ 4��� I  + 3�����
�� .prcsclicnull��� ��� uiel� 4   + /���
�� 
menI� m   - .�� ���  P r e f e r e n c e s . . .��  � 4   $ (���
�� 
menE� m   & '�� ���  O u t l o o k� 4    !���
�� 
mbri� m     �� ���  O u t l o o k� 4    ���
�� 
mbar� m    ���� � 4    ���
�� 
prcs� m    �� ��� " M i c r o s o f t   O u t l o o k� m    ���                                                                                  sevs  alis    \  Macintosh HD               �=�wBD ����System Events.app                                              �����=�w        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � ��� l  9 9��������  ��  ��  � ��� O   9 U��� O   = T��� I  D S�� ��
�� .prcsclicnull��� ��� uiel  n   D O 4   H O�
� 
butT m   K N�~�~  4   D H�}
�} 
cwin m   F G � & O u t l o o k   P r e f e r e n c e s��  � 4   = A�|
�| 
prcs m   ? @ �		 " M i c r o s o f t   O u t l o o k� m   9 :

�                                                                                  sevs  alis    \  Macintosh HD               �=�wBD ����System Events.app                                              �����=�w        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  �  l  V V�{�z�y�{  �z  �y   �x O   V� O   Z� O   c� O   l� k   u�  l  u u�w�w    -click pop up button 2    � , - c l i c k   p o p   u p   b u t t o n   2  r   u � e   u �   n   u �!"! 1   { �v
�v 
valL" 4   u {�u#
�u 
popB# m   y z�t�t  o      �s�s 0 preset Preset $%$ Z   �&'�r(& =  � �)*) o   � ��q�q 0 preset Preset* m   � �++ �,, $ 2 0 1 6   T R C   S i g n a t u r e' l  � ��p�o�n�p  �o  �n  �r  ( Z   �-.�m/- =  � �010 o   � ��l�l 0 preset Preset1 m   � �22 �33  N o n e. k   � �44 565 I  � ��k7�j
�k .prcsclicnull��� ��� uiel7 4   � ��i8
�i 
popB8 m   � ��h�h �j  6 9:9 I  � ��g;�f
�g .sysodelanull��� ��� nmbr; m   � �<< ?�      �f  : =>= l  � �?@A? I  � ��eB�d
�e .prcskprsnull���     ctxtB l  � �C�c�bC I  � ��aD�`
�a .sysontocTEXT       shorD m   � ��_�_ �`  �c  �b  �d  @   down arrow key    A �EE     d o w n   a r r o w   k e y  > FGF l  � �HIJH I  � ��^K�]
�^ .prcskprsnull���     ctxtK l  � �L�\�[L I  � ��ZM�Y
�Z .sysontocTEXT       shorM m   � ��X�X �Y  �\  �[  �]  I   down arrow key    J �NN     d o w n   a r r o w   k e y  G OPO I  � ��WQ�V
�W .sysodelanull��� ��� nmbrQ m   � �RR ?�      �V  P STS l  � �UVWU I  � ��UX�T
�U .prcskprsnull���     ctxtX l  � �Y�S�RY I  � ��QZ�P
�Q .sysontocTEXT       shorZ m   � ��O�O �P  �S  �R  �T  V  
 enter key   W �[[    e n t e r   k e yT \]\ I  � ��N^�M
�N .sysodelanull��� ��� nmbr^ m   � �__ ?�      �M  ] `�L` l  � ��K�J�I�K  �J  �I  �L  �m  / k   �aa bcb I  � ��Hd�G
�H .prcsclicnull��� ��� uield 4   � ��Fe
�F 
popBe m   � ��E�E �G  c fgf I  � ��Dh�C
�D .sysodelanull��� ��� nmbrh m   � �ii ?�      �C  g jkj l  � �lmnl I  � ��Bo�A
�B .prcskprsnull���     ctxto l  � �p�@�?p I  � ��>q�=
�> .sysontocTEXT       shorq m   � ��<�< �=  �@  �?  �A  m   down arrow key   n �rr    d o w n   a r r o w   k e yk sts I  ��;u�:
�; .sysodelanull��� ��� nmbru m   �vv ?�      �:  t wxw l yz{y I �9|�8
�9 .prcskprsnull���     ctxt| l }�7�6} I �5~�4
�5 .sysontocTEXT       shor~ m  �3�3 �4  �7  �6  �8  z  
 enter key   { �    e n t e r   k e yx ��2� I �1��0
�1 .sysodelanull��� ��� nmbr� m  �� ?�      �0  �2  % ��� r  &��� e  $�� n  $��� 1  #�/
�/ 
valL� 4  �.�
�. 
popB� m  �-�- � o      �,�, 0 preset Preset� ��+� Z  '����*�� = ',��� o  '(�)�) 0 preset Preset� m  (+�� ��� $ 2 0 1 6   T R C   S i g n a t u r e� l //�(�'�&�(  �'  �&  �*  � Z  3����%�� = 38��� o  34�$�$ 0 preset Preset� m  47�� ���  N o n e� k  ;��� ��� I ;E�#��"
�# .prcsclicnull��� ��� uiel� 4  ;A�!�
�! 
popB� m  ?@� �  �"  � ��� I FM���
� .sysodelanull��� ��� nmbr� m  FI�� ?�      �  � ��� l NY���� I NY���
� .prcskprsnull���     ctxt� l NU���� I NU���
� .sysontocTEXT       shor� m  NQ�� �  �  �  �  �   down arrow key    � ���     d o w n   a r r o w   k e y  � ��� l Ze���� I Ze���
� .prcskprsnull���     ctxt� l Za���� I Za���
� .sysontocTEXT       shor� m  Z]�� �  �  �  �  �   down arrow key    � ���     d o w n   a r r o w   k e y  � ��� I fm���
� .sysodelanull��� ��� nmbr� m  fi�� ?�      �  � ��� l nw���� I nw���
� .prcskprsnull���     ctxt� l ns���
� I ns�	��
�	 .sysontocTEXT       shor� m  no�� �  �  �
  �  �  
 enter key   � ���    e n t e r   k e y� ��� I x���
� .sysodelanull��� ��� nmbr� m  x{�� ?�      �  � ��� l ������  �  �  �  �%  � k  ���� ��� I ��� ���
�  .prcsclicnull��� ��� uiel� 4  �����
�� 
popB� m  ������ ��  � ��� I �������
�� .sysodelanull��� ��� nmbr� m  ���� ?�      ��  � ��� l ������ I �������
�� .prcskprsnull���     ctxt� l �������� I �������
�� .sysontocTEXT       shor� m  ������ ��  ��  ��  ��  �   up arrow key   � ���    u p   a r r o w   k e y� ��� I �������
�� .sysodelanull��� ��� nmbr� m  ���� ?�      ��  � ��� l ������ I �������
�� .prcskprsnull���     ctxt� l �������� I �������
�� .sysontocTEXT       shor� m  ������ ��  ��  ��  ��  �  
 enter key   � ���    e n t e r   k e y� ���� I �������
�� .sysodelanull��� ��� nmbr� m  ���� ?�      ��  ��  �+   4   l r���
�� 
sgrp� m   p q����  4   c i���
�� 
cwin� m   e h�� ���  S i g n a t u r e s 4   Z `���
�� 
prcs� m   \ _�� ��� " M i c r o s o f t   O u t l o o k m   V W���                                                                                  sevs  alis    \  Macintosh HD               �=�wBD ����System Events.app                                              �����=�w        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  �x  ��       ��������  � ��������  0 setupsignature setupSignature�� $0 updatedefaultsig updateDefaultSig
�� .aevtoappnull  �   � ****� ��n����������  0 setupsignature setupSignature��  ��  �  � �������������{�������������������������������
�� 
capp
�� kfrmID  
�� 
kocl
�� 
cSig
�� 
prdt
�� 
pnam
�� 
ctnt�� 0 	firstname  �� 0 surname  �� 0 descriptmain descriptMain�� 0 jobtitle jobTitle�� 	0 email  �� 0 phoneno phoneNo�� 0 directphone directPhone�� 
�� .corecrel****      � null�� b)���0 X*��������%�%�%�%�%�%_ %a %_ %a %_ %a %_ %a %_ %a %_ %a %_ %a %a a  UOP� ������������� $0 updatedefaultsig updateDefaultSig��  �� �����
�� 
to  �� 0 mysignature mySignature� ������
�� 
for �� 0 accountname accountName��  � �������� 0 mysignature mySignature�� 0 accountname accountName�� 0 preset Preset�  ����������������������������������+2<����������
�� .miscactvnull��� ��� null
�� 
prcs
�� 
mbar
�� 
mbri
�� 
menE
�� 
menI
�� .prcsclicnull��� ��� uiel
�� 
cwin
�� 
butT�� 
�� 
sgrp
�� 
popB
�� 
valL
�� .sysodelanull��� ��� nmbr�� 
�� .sysontocTEXT       shor
�� .prcskprsnull���     ctxt���� *j UO� **��/ "*�k/ *��/ *��/ 
*��/j UUUUUO� *��/ *��/a a /j UUO�g*�a /]*�a /S*a l/I*a l/a ,EE�O�a   hY ��a   K*a l/j Oa j Oa j j Oa j j Oa j Omj j Oa j OPY :*a l/j Oa j Oa j j Oa j Omj j Oa j O*a k/a ,EE�O�a   hY ��a   K*a k/j Oa j Oa j j Oa j j Oa j Omj j Oa j OPY :*a k/j Oa j Oa j j Oa j Omj j Oa j UUUU� �����������
�� .aevtoappnull  �   � ****� k    &��  ���  ���  �     ( 6 D R g | �		 �

 � � � � � � � � � � � I����  ��  ��  �  � =�������������� ���������������!��/��=��K��Y��l��������������������������������������������Q����
�� .sysosigtsirr   ��� null
�� 
siln�� 0 longname longName
�� 
sisn�� 0 username userName
�� 
ctxt
�� 
psof
�� 
psin�� 
�� .sysooffslong    ��� null�� 0 
rawsurname  
�� .sysoexecTEXT���     TEXT�� 0 	firstname  �� 0 surname  �� 0 	nametitle  �� 	0 email  �� 0 jobtitle jobTitle�� 0 phoneno phoneNo�� 0 directphone directPhone�� 0 address1  �� 0 	descript1  �� 0 	descript2  �� 0 logolink logoLink�� 0 weburl webURL�� 0 
weburltext 
webURLText�� 0 twitterlink twitterLink�� "0 twitterlogolink twitterLogoLink�� 0 facebooklink facebookLink�� $0 facebooklogolink facebookLogoLink�� 0 linkedinlink linkedInLink�� $0 linkedinlogolink linkedInLogoLink�� 0 	instalink 	instaLink�� 0 instalogolink instaLogoLink�� &0 bottomborderimage bottomBorderImage�� 0 descriptmain descriptMain��  0 setupsignature setupSignature��'*j   *�,E�O*�,E�UO�[�\[Zk\Z*����� 
k2E�O�j E�O�j E` Oa j E` Oa j E` Oa j E` Oa j E` Oa j E` Oa E` Oa _ %a %E` Oa  E` !Oa "E` #Oa $E` %Oa &E` 'Oa (E` )Oa *E` +Oa ,E` -Oa .E` /Oa 0E` 1Oa 2E` 3Oa 4E` 5Oa 6E` 7Oa 8E` 9O_ a :  _ !E` ;O*j+ <OPY _ E` ;O*j+ <OPascr  ��ޭ