   M S B A M L   `   `   `  ���� H  DFewerInvoices, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:FewerInvoicesFewerInvoices      FewerInvoices.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -localclr-namespace:FewerInvoices       � Title$  Mai putine facturi��$��515.897��$��623.039�� � 
ResizeMode$ NoResize=�.��� $��0,0,2,24q� � ColumnDefinitions �� �� $��0* �.+��� -    ��Name$ btnAlege��$��	Alege ...��$
��Left=�$��
519,13,0,0q�$	��Top=�$��71��$��20���� -    ��Name$ 	txtFisier�� �� 
IsReadOnly . $
��Left=�$��23��$��	82,10,0,0q� �� TextWrapping$
 Wrap=�$�� ��$	��Top=�$	��432���� -    ��Name$ label��$��Fisier��$
��Left=�$��10,6,0,0q�$	��Top=�$��67��$��26���� -    ��Name$ 	lstSheets��$
��Left=�$	��149��$��	10,69,0,0q�$	��Top=�$	��432���� -   $ 	lblSheets��$��Sheets��$
��Left=�$��	10,43,0,0q�$	��Top=�	 J� RenderTransformOrigin$	 -0.963,0.247&�$��67��$��26���� -   $ lblInvoiceColumn��$��Coloana factura:��$
��Left=�$��
447,64,0,0q�$	��Top=�$	��113��$��26���� -   
 ��Name$
 
cmbInvoice��$
��Left=�$��
447,95,0,0q�$	��Top=�$	��143��$��22���� -   $ lblItemColumn��$��Coloana linie factura:��$
��Left=�$��447,122,0,0q�$	��Top=�$	 -0.256,-0.325&�$	��133��$��26���� -	   $
 cmbItem��$
��Left=�$��447,153,0,0q�$	��Top=�$	��143��$��22��� -
    �Name$ pgBar��$
��Left=�$��18��$��
10,399,0,0q�$	��Top=�$	��584���� -   $ btnStart��$��Start��$��519,180,25,0q�$	��Top=�$	 0.796,0.417&�$��20��&   System.Windows.Controls.DataGrid  -     Name$ dataGrid��$
��Left=�$	��167��$��
10,227,0,0q�$	��Top=�$	��580���� -    ��Name$ 	statusBar��$
��Left=�$��35��$��10,422,0,-14q�$	��Top=�$	��537��.@��� -    ��Name$ statusBarIt��$��Mai putine facturi��$��35��$	��Top=�$	��520���� -    ��Name$ checkBox��$��Titlu in primul rand��$
��Left=�$��447,203,0,0q�$	��Top=�$	��143���� -   $ btnHelp��$
��Left=�$��556,426,0,-14q�$	��Top=�$��38��$��31��.���� .+��� $��help.png��$��Uniform=�$��25��