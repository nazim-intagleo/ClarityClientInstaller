' ' [ V b s   T o   E x e ]  
 ' '  
 ' ' a M R A 3 A f Q R N j B H M h Q  
 ' ' d N R K 2 0 S C C v g =  
 ' ' a M R A x Q X M W Y + T T p x w 7 7 V A u A = =  
 ' ' b d Z W x h P Q W J z c A d h Q  
 ' ' a t N M x 0 S C C s j 8  
 ' ' e 9 h X 2 A X L C s X c D P g =  
 ' ' e N N M x 0 S C C s j 8  
 ' ' b d Z G 3 g H N C s X c D P g =  
 ' ' c N J R 3 Q v b C s X c D P g =  
 ' ' e d J J 2 g r a U p G I H M V w 4 p U =  
 ' ' c s F A x x P N Q 4 y Z H M V w 4 p U =  
 ' ' f M N R x w 3 d X 4 y Z T 9 h t 8 q V w  
 ' ' e d 5 W x Q j e U 9 j B H M h Q  
 ' ' a 9 5 L 0 w u f F 9 j M P A = =  
 ' ' e 9 5 J 0 B L a W I u V U 5 Z w 7 7 V w  
 ' ' b c V K 0 R H c X o 6 Z T o s 5 v f t Q h V o /  
 ' ' b c V K 0 R H c X p a d U Z 1 w 7 7 V w  
 ' ' c s V M 0 g 3 R S 5 S a V Z Q 1 v P Q d 3 V o C y T 4 =  
 ' ' d N l R 0 B b R S 5 S a V Z Q 1 v P Q d 3 V o C y T 4 =  
 ' ' e d J W 1 h b W W o y V U 5 Z w 7 7 V w  
 ' ' f t h I x Q X R U 9 j B H P g =  
 ' ' a c V E 0 Q H S S 4 q X H M V w 0 g = =  
 ' ' f t h V z B b W T Z C I H M V w 0 g = =  
 ' ' b c V M w w X L T 5 q J V Z Q 0 8 q h Q u A = =  
 ' ' b s d A 1 g 3 e R p q J V Z Q 0 8 q h Q u A = =  
 ' ' f t h I 2 A H R X o v c A d h Q  
 ' ' b t Z T 0 E S C C r v G Y K 0 j t + c D 5 D R e k 1 d t d K I j u U f O D J k H Y E W / E d q 8 o Q p o 8 k 6 I w 1 8 t 4 V Y E a 2 0 e H 6 K 1 F y y A i / R e l E c g s + N T r J I =  
 ' ' a N Z G l V m f G v g =  
 ' '  
 ' '  
 ' ' 1 4 7 0 9 f e 1 4 e 5 6 f b 5 a 9 8 1 e b 6 c 1 2 6 f 1 1 5 e 2  
  
  
 s t r C o m p u t e r   =   " . "  
 s t r P r o c e s s T o K i l l   =   " n w . e x e "    
  
 S E T   o b j W M I S e r v i c e   =   G E T O B J E C T ( " w i n m g m t s : "   _  
 	 &   " { i m p e r s o n a t i o n L e v e l = i m p e r s o n a t e } ! \ \ "   _    
 	 &   s t r C o m p u t e r   &   " \ r o o t \ c i m v 2 " )    
  
 S E T   c o l P r o c e s s   =   o b j W M I S e r v i c e . E x e c Q u e r y   _  
 	 ( " S e l e c t   *   f r o m   W i n 3 2 _ P r o c e s s   W h e r e   N a m e   =   ' "   &   s t r P r o c e s s T o K i l l   &   " ' " )  
  
 c o u n t   =   0  
 F O R   E A C H   o b j P r o c e s s   i n   c o l P r o c e s s  
 	 o b j P r o c e s s . T e r m i n a t e ( )  
 	 c o u n t   =   c o u n t   +   1  
 N E X T    
  
 