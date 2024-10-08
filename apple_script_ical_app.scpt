FasdUAS 1.101.10   ��   ��    k             l     ��  ��    A ; Prompt the user to select the folder containing .ics files     � 	 	 v   P r o m p t   t h e   u s e r   t o   s e l e c t   t h e   f o l d e r   c o n t a i n i n g   . i c s   f i l e s   
  
 l    	 ����  r     	    I    ���� 
�� .sysostflalis    ��� null��    �� ��
�� 
prmp  m       �   X S e l e c t   t h e   f o l d e r   c o n t a i n i n g   y o u r   . i c s   f i l e s��    o      ���� 0 	icsfolder 	icsFolder��  ��        l     ��������  ��  ��        l     ��  ��    / ) Get the list of .ics files in the folder     �   R   G e t   t h e   l i s t   o f   . i c s   f i l e s   i n   t h e   f o l d e r      l  
  ����  O   
     r         6    ! " ! n     # $ # 2    ��
�� 
file $ o    ���� 0 	icsfolder 	icsFolder " =    % & % 1    ��
�� 
nmxt & m     ' ' � ( (  i c s   o      ���� 0 icsfiles icsFiles  m   
  ) )�                                                                                  MACS  alis    @  Macintosh HD               ����BD ����
Finder.app                                                     ��������        ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  ��  ��     * + * l     ��������  ��  ��   +  , - , l     �� . /��   . C = Loop through each .ics file and create a new calendar for it    / � 0 0 z   L o o p   t h r o u g h   e a c h   . i c s   f i l e   a n d   c r e a t e   a   n e w   c a l e n d a r   f o r   i t -  1�� 1 l   � 2���� 2 O    � 3 4 3 X   " � 5�� 6 5 k   2 � 7 7  8 9 8 l  2 2�� : ;��   : A ; Get the name of the .ics file (without the .ics extension)    ; � < < v   G e t   t h e   n a m e   o f   t h e   . i c s   f i l e   ( w i t h o u t   t h e   . i c s   e x t e n s i o n ) 9  = > = r   2 7 ? @ ? n   2 5 A B A 1   3 5��
�� 
pnam B o   2 3���� 0 icsfile icsFile @ o      ���� 0 icsname icsName >  C D C l  8 K E F G E r   8 K H I H n   8 G J K J 7  9 G�� L M
�� 
ctxt L m   ? A����  M m   B F������ K o   8 9���� 0 icsname icsName I o      ���� 0 calendarname calendarName F . ( Remove .ics extension for calendar name    G � N N P   R e m o v e   . i c s   e x t e n s i o n   f o r   c a l e n d a r   n a m e D  O P O l  L L��������  ��  ��   P  Q R Q l  L L�� S T��   S ; 5 Create a new calendar with the name of the .ics file    T � U U j   C r e a t e   a   n e w   c a l e n d a r   w i t h   t h e   n a m e   o f   t h e   . i c s   f i l e R  V W V r   L d X Y X I  L `���� Z
�� .corecrel****      � null��   Z �� [ \
�� 
kocl [ m   N Q��
�� 
wres \ �� ]��
�� 
prdt ] K   T Z ^ ^ �� _��
�� 
pnam _ o   U X���� 0 calendarname calendarName��  ��   Y o      ���� 0 newcalendar newCalendar W  ` a ` l  e e��������  ��  ��   a  b c b l  e e�� d e��   d * $ Get the POSIX path of the .ics file    e � f f H   G e t   t h e   P O S I X   p a t h   o f   t h e   . i c s   f i l e c  g h g r   e r i j i n   e n k l k 1   j n��
�� 
psxp l l  e j m���� m c   e j n o n o   e f���� 0 icsfile icsFile o m   f i��
�� 
alis��  ��   j o      ���� 0 icsfilepath icsFilePath h  p q p l  s s��������  ��  ��   q  r s r l  s s�� t u��   t L F Use the `do shell script` command to open the .ics file with Calendar    u � v v �   U s e   t h e   ` d o   s h e l l   s c r i p t `   c o m m a n d   t o   o p e n   t h e   . i c s   f i l e   w i t h   C a l e n d a r s  w x w I  s ��� y��
�� .sysoexecTEXT���     TEXT y b   s ~ z { z m   s v | | � } } " o p e n   - a   C a l e n d a r   { n   v } ~  ~ 1   y }��
�� 
strq  o   v y���� 0 icsfilepath icsFilePath��   x  ��� � l  � ���������  ��  ��  ��  �� 0 icsfile icsFile 6 o   % &���� 0 icsfiles icsFiles 4 m     � ��                                                                                  wrbt  alis    8  Macintosh HD               ����BD ����Calendar.app                                                   ��������        ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  ��  ��  ��       �� � ���   � ��
�� .aevtoappnull  �   � **** � �� ����� � ���
�� .aevtoappnull  �   � **** � k     � � �  
 � �   � �  1����  ��  ��   � ���� 0 icsfile icsFile � �� ���� )�� ��� '�� ��������������������������������� |����
�� 
prmp
�� .sysostflalis    ��� null�� 0 	icsfolder 	icsFolder
�� 
file �  
�� 
nmxt�� 0 icsfiles icsFiles
�� 
kocl
�� 
cobj
�� .corecnte****       ****
�� 
pnam�� 0 icsname icsName
�� 
ctxt������ 0 calendarname calendarName
�� 
wres
�� 
prdt�� 
�� .corecrel****      � null�� 0 newcalendar newCalendar
�� 
alis
�� 
psxp�� 0 icsfilepath icsFilePath
�� 
strq
�� .sysoexecTEXT���     TEXT�� �*��l E�O� ��-�[�,\Z�81E�UO� i f�[��l kh  ��,E�O�[a \[Zk\Za 2E` O*�a a �_ la  E` O�a &a ,E` Oa _ a ,%j OP[OY��Uascr  ��ޭ