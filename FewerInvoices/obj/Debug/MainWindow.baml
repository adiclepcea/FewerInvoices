   M S B A M L   `   `   `  ����H  DFewerInvoices, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:FewerInvoicesFewerInvoices  5   	       FewerInvoices.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:FewerInvoices     5   	     � Title$  Mai putine facturi��5   	   $��515.897��6$   $��623.039��65    � 
ResizeMode$ NoResize=�6E   .��5	      � $��0,0,2,24q�6    � ColumnDefinitions 5
   
   �� 5      �� 5      $��0* �6   .+�5   
   �� -   5         ��Name$ btnAlege��5      $��	Alege ...��6#   $
��Left=�67   $��
519,13,0,0q�6R   $	��Top=�6f   $��71��6~   $��20��6�   �� 5   
   -   5         ��Name$ 	txtFisier��5       �� 
IsReadOnly . $
��Left=�67   $��23��6R   $��	82,10,0,0q�6^    �� TextWrapping$
 Wrap=�6q   $�� ��6�   $	��Top=�6�   $	��432��6�   �� 5   
   -   5         ��Name$ label��5      $��Fisier��6   $
��Left=�60   $��10,6,0,0q�6K   $	��Top=�6]   $��67��6u   $��26��6�   �� 5   
   -   5         ��Name$ 	lstSheets��5      $
��Left=�6%   $	��149��6@   $��	10,69,0,0q�6M   $	��Top=�6`   $	��432��6x   6�   �� 5   
   -   5        $ 	lblSheets��5      $��Sheets��6#   $
��Left=�64   $��	10,43,0,0q�6O   $	��Top=�6b   	 J� RenderTransformOrigin$	 -0.963,0.247&�6z   $��67��6�   $��26��6�   �� 5   
   -   5        $ lblInvoiceColumn��5      $��Coloana factura:��6*   $
��Left=�6E   $��
447,64,0,0q�6`   $	��Top=�6t   $	��113��6�   $��26��6�   �� 5   
   -   5        
 ��Name$
 
cmbInvoice��5      $
��Left=�6'   $��
447,95,0,0q�6B   $	��Top=�6V   $	��143��6n   $��22��6z   �� 5   
   -   5        $ lblItemColumn��5      $��Coloana linie factura:��6'   $
��Left=�6H   $��447,122,0,0q�6c   $	��Top=�6x   $	 -0.256,-0.325&�6�   $	��133��6�   $��26��6�   �� 5   
   -	   5        $
 cmbItem��5      $
��Left=�6$   $��447,153,0,0q�6?   $	��Top=�6T   $	��143��6l   $��22��6x   � 5   
   -
   5         �Name$ pgBar��5      $
��Left=�6%   $��18��6@   $��
10,399,0,0q�6L   $	��Top=�6`   $	��584��6x   �� 5   
   -   5        $ btnStart��5      $��Start��6#   $��519,180,25,0q�63   $	��Top=�6I   $	 0.796,0.417&�6a   $��20��6�   6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ dataGrid��5      $
��Left=�6%   $	��167��6@   $��
10,227,0,0q�6M   $	��Top=�6a   $	��580��6y   �� 5   
   -   5         ��Name$ 	statusBar��5      $
��Left=�6'   $��35��6B   $��10,422,0,-14q�6N   $	��Top=�6d   $	��537��6|   .@�5      �� -   5         ��Name$ statusBarIt��5      $��Mai putine facturi��61   $��35��6N   $	��Top=�6Z   $	��520��6r   5      �� 5   
   -   5         ��Name$ checkBox��5      $��Titlu in primul rand��6%   $
��Left=�6D   $��447,203,0,0q�6_   $	��Top=�6t   $	��143��6�   �� 5   
   -   5        $ btnHelp��5      $
��Left=�6"   $��556,426,0,-14q�6=   $	��Top=�6T   $��38��6l   $��31��6w   .��5      �� .+�5       �� $��help.png��6   $��Uniform=�6*   $��25��6<   5!      5"      5$      5%      