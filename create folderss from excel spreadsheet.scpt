FasdUAS 1.101.10   ��   ��    k             l        	  x     �� 
 ��   
 1      ��
�� 
ascr  �� ��
�� 
minv  m         �    2 . 4��       Yosemite (10.10) or later    	 �   4   Y o s e m i t e   ( 1 0 . 1 0 )   o r   l a t e r      x    �� ����    2  	 ��
�� 
osax��        l     ��������  ��  ��        l     ��  ��    T Ndefault choose file dialog. THe type is an Apple UTI, uniform type identifier,     �   � d e f a u l t   c h o o s e   f i l e   d i a l o g .   T H e   t y p e   i s   a n   A p p l e   U T I ,   u n i f o r m   t y p e   i d e n t i f i e r ,      l     ��  ��    ? 9not the other kind. returns an alias to the file you pick     �   r n o t   t h e   o t h e r   k i n d .   r e t u r n s   a n   a l i a s   t o   t h e   f i l e   y o u   p i c k       l     !���� ! r      " # " I    ���� $
�� .sysostdfalis    ��� null��   $ �� % &
�� 
prmp % l 	   '���� ' m     ( ( � ) ) V C h o o s e   t h e   E x c e l   f i l e   w i t h   t h e   f o l d e r   n a m e s��  ��   & �� * +
�� 
dflc * l   	 ,���� , I   	�� -��
�� .earsffdralis        afdr - m    ��
�� afdrdesk��  ��  ��   + �� . /
�� 
ftyp . J   
  0 0  1�� 1 m   
  2 2 � 3 3 L o r g . o p e n x m l f o r m a t s . s p r e a d s h e e t m l . s h e e t��   / �� 4 5
�� 
lfiv 4 m    ��
�� boovfals 5 �� 4��
�� 
mlsl��   # o      ����  0 theexcelsource theExcelSource��  ��      6 7 6 l     ��������  ��  ��   7  8 9 8 l   � :���� : O    � ; < ; k    � = =  > ? > l   �� @ A��   @ \ Vyes, it's "open workbook" the command and "workbook file name (text version of alias)"    A � B B � y e s ,   i t ' s   " o p e n   w o r k b o o k "   t h e   c o m m a n d   a n d   " w o r k b o o k   f i l e   n a m e   ( t e x t   v e r s i o n   o f   a l i a s ) " ?  C D C l   �� E F��   E  this will bite you    F � G G $ t h i s   w i l l   b i t e   y o u D  H I H r    ) J K J I   %���� L
�� .smXL1169null��� ��� null��   L �� M��
�� 
WbFN M l   ! N���� N c    ! O P O o    ����  0 theexcelsource theExcelSource P m     ��
�� 
ctxt��  ��  ��   K o      ���� .0 convertexceltofolders convertExcelToFolders I  Q R Q l  * *��������  ��  ��   R  S T S l  * *�� U V��   U  get the active worksheet    V � W W 0 g e t   t h e   a c t i v e   w o r k s h e e t T  X Y X r   * 5 Z [ Z n   * 1 \ ] \ 1   - 1��
�� 
1107 ] o   * -���� .0 convertexceltofolders convertExcelToFolders [ o      ���� 80 convertexceltofolderssheet convertExcelToFoldersSheet Y  ^ _ ^ l  6 6�� ` a��   `  get the used range    a � b b $ g e t   t h e   u s e d   r a n g e _  c d c r   6 A e f e n   6 = g h g 1   9 =��
�� 
1756 h o   6 9���� 80 convertexceltofolderssheet convertExcelToFoldersSheet f o      ���� $0 thenamelistrange theNameListRange d  i j i l  B B��������  ��  ��   j  k l k l  B B�� m n��   m V Pcount the rows in the range, a more simplistic way to do it, but not as specific    n � o o � c o u n t   t h e   r o w s   i n   t h e   r a n g e ,   a   m o r e   s i m p l i s t i c   w a y   t o   d o   i t ,   b u t   n o t   a s   s p e c i f i c l  p q p r   B Q r s r l  B M t���� t I  B M�� u��
�� .corecnte****       **** u n   B I v w v 2  E I��
�� 
crow w o   B E���� $0 thenamelistrange theNameListRange��  ��  ��   s o      ���� 0 numcells numCells q  x y x l  R R�� z {��   z  build our end cell lable    { � | | 0 b u i l d   o u r   e n d   c e l l   l a b l e y  } ~ } r   R ]  �  b   R Y � � � m   R U � � � � �  A � o   U X���� 0 numcells numCells � o      ���� 0 endcell endCell ~  � � � l  ^ ^�� � ���   � k eapplescript has a number of built-in UI primitives that don't require as much work as powershell does    � � � � � a p p l e s c r i p t   h a s   a   n u m b e r   o f   b u i l t - i n   U I   p r i m i t i v e s   t h a t   d o n ' t   r e q u i r e   a s   m u c h   w o r k   a s   p o w e r s h e l l   d o e s �  � � � l  ^ ^�� � ���   � G Athis creates an alias to where you want the folders to be created    � � � � � t h i s   c r e a t e s   a n   a l i a s   t o   w h e r e   y o u   w a n t   t h e   f o l d e r s   t o   b e   c r e a t e d �  � � � l  ^ ^�� � ���   � Z Tyou get the "create new folder" button for free in the dialog, no need to specify it    � � � � � y o u   g e t   t h e   " c r e a t e   n e w   f o l d e r "   b u t t o n   f o r   f r e e   i n   t h e   d i a l o g ,   n o   n e e d   t o   s p e c i f y   i t �  � � � r   ^ s � � � I  ^ o���� �
�� .sysostflalis    ��� null��   � �� � �
�� 
prmp � l 	 ` c ����� � m   ` c � � � � � ^ S e l e c t   w h e r e   y o u   w a n t   t h e   f o l d e r s   t o   b e   c r e a t e d��  ��   � �� ���
�� 
dflc � l  d i ����� � I  d i�� ���
�� .earsffdralis        afdr � m   d e��
�� afdrdesk��  ��  ��  ��   � o      ���� &0 destinationfolder destinationFolder �  � � � l  t t��������  ��  ��   �  � � � l  t t�� � ���   � j dset up our beginning and ending of the range. this is effectively the same as the powershell version    � � � � � s e t   u p   o u r   b e g i n n i n g   a n d   e n d i n g   o f   t h e   r a n g e .   t h i s   i s   e f f e c t i v e l y   t h e   s a m e   a s   t h e   p o w e r s h e l l   v e r s i o n �  � � � l  t t�� � ���   �   just a bit more simplistic    � � � � 4 j u s t   a   b i t   m o r e   s i m p l i s t i c �  � � � r   t  � � � b   t { � � � m   t w � � � � �  A 1 : � o   w z���� 0 endcell endCell � o      ���� 0 therange theRange �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � @ :get an (applescript) list of cells in the range from Excel    � � � � t g e t   a n   ( a p p l e s c r i p t )   l i s t   o f   c e l l s   i n   t h e   r a n g e   f r o m   E x c e l �  � � � r   � � � � � n  � � � � � 2   � ���
�� 
ccel � n   � � � � � 4   � ��� �
�� 
X117 � o   � ����� 0 therange theRange � o   � ����� 80 convertexceltofolderssheet convertExcelToFoldersSheet � o      ���� 0 thenamecells theNameCells �  � � � l  � ��� � ���   � ' !iterate through the list of cells    � � � � B i t e r a t e   t h r o u g h   t h e   l i s t   o f   c e l l s �  � � � X   � � ��� � � k   � � � �  � � � l  � ��� � ���   � % pull value2 for the folder name    � � � � > p u l l   v a l u e 2   f o r   t h e   f o l d e r   n a m e �  � � � r   � � � � � n   � � � � � 1   � ���
�� 
DPV2 � o   � ����� 0 thecell theCell � o      ���� 0 
foldername   �  � � � l  � ��� � ���   � d ^we have to explicitly target the finder with this, a convention dating back to the early 1990s    � � � � � w e   h a v e   t o   e x p l i c i t l y   t a r g e t   t h e   f i n d e r   w i t h   t h i s ,   a   c o n v e n t i o n   d a t i n g   b a c k   t o   t h e   e a r l y   1 9 9 0 s �  � � � l  � ��� � ���   � 6 0a lot of folder/file stuff is part of the finder    � � � � ` a   l o t   o f   f o l d e r / f i l e   s t u f f   i s   p a r t   o f   t h e   f i n d e r �  ��� � O   � � � � � k   � � � �  � � � l  � ��� � ���   � @ :make a new folder in the destination with the desired name    � � � � t m a k e   a   n e w   f o l d e r   i n   t h e   d e s t i n a t i o n   w i t h   t h e   d e s i r e d   n a m e �  ��� � I  � ����� �
�� .corecrel****      � null��   � �� � �
�� 
kocl � m   � ���
�� 
cfol � �� � �
�� 
insh � o   � ����� &0 destinationfolder destinationFolder � �� ���
�� 
prdt � K   � � � � �� ���
�� 
pnam � o   � ����� 0 
foldername  ��  ��  ��   � m   � � � ��                                                                                  MACS  alis    4  Lancer                     ߶�=BD ����
Finder.app                                                     ����߶�=        ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p    L a n c e r  &System/Library/CoreServices/Finder.app  / ��  ��  �� 0 thecell theCell � o   � ��� 0 thenamecells theNameCells �  � � � l  � ��~ � ��~   � : 4we're done, quit the app because we're nice that way    � � � � h w e ' r e   d o n e ,   q u i t   t h e   a p p   b e c a u s e   w e ' r e   n i c e   t h a t   w a y �  ��} � I  � ��|�{�z
�| .aevtquitnull��� ��� null�{  �z  �}   < m     � ��                                                                                  XCEL  alis    :  Lancer                     ߶�=BD ����Microsoft Excel.app                                            �����7�         ����  
 cu             Applications  #/:Applications:Microsoft Excel.app/   (  M i c r o s o f t   E x c e l . a p p    L a n c e r   Applications/Microsoft Excel.app  / ��  ��  ��   9  ��y � l     �x�w�v�x  �w  �v  �y       �u � �u   � �t�s
�t 
pimr
�s .aevtoappnull  �   � ****  �r�r    �q �p
�q 
vers�p   �o�n
�o 
cobj    �m
�m 
osax�n   �l�k�j	�i
�l .aevtoappnull  �   � **** k     �

    8�h�h  �k  �j   �g�g 0 thecell theCell	 /�f (�e�d�c�b 2�a�`�_�^�] ��\�[�Z�Y�X�W�V�U�T�S�R ��Q ��P�O�N ��M�L�K�J�I�H�G�F ��E�D�C�B�A�@�?
�f 
prmp
�e 
dflc
�d afdrdesk
�c .earsffdralis        afdr
�b 
ftyp
�a 
lfiv
�` 
mlsl�_ 

�^ .sysostdfalis    ��� null�]  0 theexcelsource theExcelSource
�\ 
WbFN
�[ 
ctxt
�Z .smXL1169null��� ��� null�Y .0 convertexceltofolders convertExcelToFolders
�X 
1107�W 80 convertexceltofolderssheet convertExcelToFoldersSheet
�V 
1756�U $0 thenamelistrange theNameListRange
�T 
crow
�S .corecnte****       ****�R 0 numcells numCells�Q 0 endcell endCell�P 
�O .sysostflalis    ��� null�N &0 destinationfolder destinationFolder�M 0 therange theRange
�L 
X117
�K 
ccel�J 0 thenamecells theNameCells
�I 
kocl
�H 
cobj
�G 
DPV2�F 0 
foldername  
�E 
cfol
�D 
insh
�C 
prdt
�B 
pnam�A 
�@ .corecrel****      � null
�? .aevtquitnull��� ��� null�i �*����j ��kv�f�f� 
E�O� �*���&l E` O_ a ,E` O_ a ,E` O_ a -j E` Oa _ %E` O*�a ��j a  E` Oa _ %E` O_ a  _ /a !-E` "O I_ "[a #a $l kh  �a %,E` &Oa '  *a #a (a )_ a *a +_ &la , -U[OY��O*j .Uascr  ��ޭ