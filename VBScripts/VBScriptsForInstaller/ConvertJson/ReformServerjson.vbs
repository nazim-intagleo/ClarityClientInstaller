 
  
 s t r I n p u t   =   U s e r I n p u t (   " E n t e r   e n v i r o n m e n t [ D e v ,   Q e ,   D e m o ,   P r o d u c t i o n ] "   )  
  
 i f   L C a s e ( s t r I n p u t )   =   " d e v "   t h e n  
 S e r v e r D a t a K e y   =   " s e r v e r D a t a "  
 S e r v e r D a t a C h a n g e   =   " h t t p : / / c l a r i t y d a t a . a g s . c o m / d e v c l a r i t y d a t a / "  
 F o n t S e r v e r K e y   =   " f o n t S e r v e r "  
 F o n t S e r v e r C h a n g e   =   " h t t p : / / c l a r i t y d a t a . a g s . c o m / r e s o u r c e s / D e v / "  
  
 E l s e I f   L C a s e ( s t r I n p u t ) = " q e "   t h e n  
 S e r v e r D a t a K e y   =   " s e r v e r D a t a "  
 S e r v e r D a t a C h a n g e   =   " h t t p : / / c l a r i t y d a t a . a g s . c o m / q e c l a r i t y d a t a / "  
 F o n t S e r v e r K e y   =   " f o n t S e r v e r "  
 F o n t S e r v e r C h a n g e   =   " h t t p : / / c l a r i t y d a t a . a g s . c o m / r e s o u r c e s / Q e / "  
  
 E l s e I f   L C a s e ( s t r I n p u t ) = " d e m o "   t h e n  
 S e r v e r D a t a K e y   =   " s e r v e r D a t a "  
 S e r v e r D a t a C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / d e m o c l a r i t y d a t a / "  
 F o n t S e r v e r K e y   =   " f o n t S e r v e r "  
 F o n t S e r v e r C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / r e s o u r c e s / D e m o / "  
  
 E l s e I f   L C a s e ( s t r I n p u t ) = " p r o d u c t i o n "   t h e n  
 S e r v e r D a t a K e y   =   " s e r v e r D a t a "  
 S e r v e r D a t a C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / c l a r i t y d a t a / "  
 F o n t S e r v e r K e y   =   " f o n t S e r v e r "  
 F o n t S e r v e r C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / r e s o u r c e s / P r o d / "  
  
 e l s e  
 S e r v e r D a t a K e y   =   " s e r v e r D a t a "  
 S e r v e r D a t a C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / c l a r i t y d a t a / "  
 F o n t S e r v e r K e y   =   " f o n t S e r v e r "  
 F o n t S e r v e r C h a n g e   =   " h t t p s : / / c l a r i t y d a t a . a l l u r e g l o b a l . c o m / r e s o u r c e s / P r o d / "  
  
 E n d   I f  
  
 S e t   o b j F S O   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
 S e t   o b j F i l e   =   o b j F S O . O p e n T e x t F i l e (   " C : \ C l a r i t y C l i e n t I n s t a l l e r \ s e r v e r s . j s o n " ,   1   )  
 R e a d   =   	 o b j F i l e . R e a d A l l  
  
 S e t   o R E   =   N e w   R e g e x p  
 o R E . G l o b a l   =   T r u e  
 o R E . I g n o r e C a s e   =   F a l s e  
 o R E . P a t t e r n   =   " " " "   &   S e r v e r D a t a K e y   & " " " : \ s " " ( . + ? ) " " "  
 S e t   c o l M a t c h e s   =   o R E . E x e c u t e   (   R e a d   )  
 	  
 F o r   E a c h   o M a t c h   i n   c o l M a t c h e s  
         M a t c h   =   o M a t c h . S u b M a t c h e s ( 0 )  
 N e x t  
         W r i t e   =   R e p l a c e   (   R e a d ,   M a t c h ,   S e r v e r D a t a C h a n g e ,   1 ,   - 1 ,   0   )  
  
 S e t   o b j F i l e   =   o b j F S O . O p e n T e x t F i l e (   " C : \ C l a r i t y C l i e n t I n s t a l l e r \ s e r v e r s . j s o n " ,   2   )  
 o b j F i l e . W r i t e   W r i t e  
 o b j F i l e . C l o s e  
  
 '   U p d a t e   f o n t S e r v e r   k e t   o f   S e r v e r . j s o n   f i l e  
 S e t   o b j F i l e   =   o b j F S O . O p e n T e x t F i l e (   " C : \ C l a r i t y C l i e n t I n s t a l l e r \ s e r v e r s . j s o n " ,   1   )  
 R e a d   =   	 o b j F i l e . R e a d A l l  
  
   S e t   o R E   =   N e w   R e g e x p  
   o R E . G l o b a l   =   T r u e  
   o R E . I g n o r e C a s e   =   F a l s e  
   o R E . P a t t e r n   =   " " " "   &   F o n t S e r v e r K e y   & " " " : \ s " " ( . + ? ) " " "  
   S e t   c o l M a t c h e s   =   o R E . E x e c u t e   (   R e a d   )  
 	  
   F o r   E a c h   o M a t c h   i n   c o l M a t c h e s  
           M a t c h   =   o M a t c h . S u b M a t c h e s ( 0 )  
   N e x t  
           W r i t e   =   R e p l a c e   (   R e a d ,   M a t c h ,   F o n t S e r v e r C h a n g e ,   1 ,   - 1 ,   0   )  
  
 S e t   o b j F i l e   =   o b j F S O . O p e n T e x t F i l e (   " C : \ C l a r i t y C l i e n t I n s t a l l e r \ s e r v e r s . j s o n " ,   2   )  
 o b j F i l e . W r i t e   W r i t e  
 o b j F i l e . C l o s e  
  
  
  
 F u n c t i o n   U s e r I n p u t (   m y P r o m p t   )  
 '   T h i s   f u n c t i o n   p r o m p t s   t h e   u s e r   f o r   s o m e   i n p u t .  
 '   W h e n   t h e   s c r i p t   r u n s   i n   C S C R I P T . E X E ,   S t d I n   i s   u s e d ,  
 '   o t h e r w i s e   t h e   V B S c r i p t   I n p u t B o x (   )   f u n c t i o n   i s   u s e d .  
 '   m y P r o m p t   i s   t h e   t h e   t e x t   u s e d   t o   p r o m p t   t h e   u s e r   f o r   i n p u t .  
 '   T h e   f u n c t i o n   r e t u r n s   t h e   i n p u t   t y p e d   e i t h e r   o n   S t d I n   o r   i n   I n p u t B o x (   ) .  
 '   W r i t t e n   b y   R o b   v a n   d e r   W o u d e  
 '   h t t p : / / w w w . r o b v a n d e r w o u d e . c o m  
         '   C h e c k   i f   t h e   s c r i p t   r u n s   i n   C S C R I P T . E X E  
         I f   U C a s e (   R i g h t (   W S c r i p t . F u l l N a m e ,   1 2   )   )   =   " \ C S C R I P T . E X E "   T h e n  
                 '   I f   s o ,   u s e   S t d I n   a n d   S t d O u t  
                 W S c r i p t . S t d O u t . W r i t e   m y P r o m p t   &   "   "  
                 U s e r I n p u t   =   W S c r i p t . S t d I n . R e a d L i n e  
         E l s e  
                 '   I f   n o t ,   u s e   I n p u t B o x (   )  
                 U s e r I n p u t   =   I n p u t B o x (   m y P r o m p t   )  
         E n d   I f  
 E n d   F u n c t i o n  
  
 