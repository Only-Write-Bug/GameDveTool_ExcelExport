   M S B A M L   `   `   `  ����J  FExcelExportTool, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:ExcelExportToolExcelExportTool  5   	        ExcelExportTool.MainWindow   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:ExcelExportTool     5   	     � Title$  Excel Export Tool��5   	   $	��450��6#   $	��800��60   
���iii� � 
ResizeMode$ NoResize=�5
   	   .��5      � .+�5   
   �� $	��Top=�5      $
��Left=�5      $��	93,63,0,0q�5      Excel folder path:5      �� 5   
   -   5         ��Name$ ExcelPathTextBox��5      $	��400��5      $��20��5      $	��Top=�5      $��0,60,0,0q�5      $��15(�5      .��6   5      �� 5   
   -   5         ��Name$ ExcelPathEditorButton��5      $��50��5      $��20��5      $��Right=�5       $	��Top=�5!      $��
0,60,147,0q�5"      $��Editor��5#      6    �� 5%   
   $
��Left=�5&      $	��Top=�5'      $��	93,90,0,0q�5(      XML folder path:5*      �� 5+   
   -   5        $ XMLPathTextBox��5,      $	��400��5-      $��20��5.      $	��Top=�5/      $��0,90,0,0q�50      $��15(�52      .��6   53      �� 54   
   -   5        $ XmlPathEditorButton��55      $��50��56      $��20��57      $��Right=�58      $	��Top=�59      $��
0,90,147,0q�5:      $��Editor��5;      6    �� 5=   
   -   5        $ StartButton��5>      $	��200��5?      $��50��5@      $��Bottom=�5A      $��0,0,0,30q�5B      $��Start��5D      $��35(�5E      5F      5G      5H      