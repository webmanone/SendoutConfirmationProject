   M S B A M L   `   `   `  ����Z  VSendout Calendar Invite Project, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullQ-clr-namespace:Sendout_Calendar_Invite_ProjectSendout_Calendar_Invite_Project  5   	   0    *Sendout_Calendar_Invite_Project.MainWindow   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ] YXceed.Wpf.Toolkit.NET5, Version=4.5.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f43xct)http://schemas.xceed.com/wpf/xaml/toolkit  5   	   9local-clr-namespace:Sendout_Calendar_Invite_Project  5   	     � Title$)  #Sendout Confirmation Invite Creator��5	   	   $	��956��65   $	��374��6A   .��5
      � $��10q�6   $��Stretch=�6   $��Stretch=�65    � RowDefinitions 5   
   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��* �6    � ColumnDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5      $
��Auto �6   �� 5      $��* �6   .+�5   
   �� $��0��6   $��0��6   $��	Template:��6-   $��Center=�6A   �� 5   
   -   5         ��Name$ TemplateComboBox��5      $��0��6-   $��1��6:   $	��200��6J   $��5q�6V    �� SelectedIndex$ 0��6a   $��Center=�6s   .@�5      �� .��6   First stage phone call63   �� 5      .��6   Teams interview6,   �� 5       .��6   In-person interview60   �� 5!      .��6   Other6"   5"      �� 5$   
   $��0��6   $��2��6   $��
Date/time:��6-   $��Center=�6B   &   Xceed.Wpf.Toolkit.DateTimePicker  5%   
   -   5          Name$ DateTimePicker��5%      $��0��65   $��3��6B   $	��200��6R   $��5q�6^   6i   �� 5'   
   $��1��6   $��0��6   $��Client time zone:��6-   $��Center=�6I   �� 5(   
   -   5        $ ClientComboBox��5(      $��1��6+   $��1��68   $	��200��6H   $��5q�6T   $ 0��6_   $��Center=�6q   .@�5)      �� .��6   	Eastern6$   �� 5*      .��6   	Central6$   �� 5+      .��6   
Mountain6%   �� 5,      .��6   	Pacific6$   5-      �� 5/   
   $��1��6   $��2��6   $��Candidate time zone:��6-   $��Center=�6L   �� 50   
   -   5        $ CandidateComboBox��50      $��1��6.   $��4��6;   $	��200��6K   $��5q�6W   $ 0��6b   $��Center=�6t   .@�51      �� .��6   	Eastern6$   �� 52      .��6   	Central6$   �� 53      .��6   
Mountain6%   �� 54      .��6   	Pacific6$   55      �� 57   
   $��2��6   $��0��6   $��Client name:��6-   $��Center=�6D   �� 58   
   -   5         ��Name$ ClientNameTextBox��58      $��2��6-   $��1��6:   $	��200��6J   $��5q�6V   �� 5:   
   $��2��6   $��2��6   $��Candidate name:��6-   $��Center=�6G   �� 5;   
   -   5        $ CandidateNameTextBox��5;      $��2��60   $��3��6=   $	��200��6M   $��5q�6Y   �� 5=   
   $��3��6   $��0��6   $��Client email:��6-   $��Center=�6E   �� 5>   
   -   5        $ ClientEmailTextBox��5>      $��3��6.   $��1��6;   $	��200��6K   $��5q�6W   �� 5@   
   $��3��6   $��2��6   $��Candidate email:��6-   $��Center=�6H   �� 5A   
   -   5        $ CandidateEmailTextBox��5A      $��3��61   $��3��6>   $	��200��6N   $��5q�6Z   �� 5C   
   $��4��6   $��0��6   $��Client company:��6-   $��Center=�6G   �� 5D   
   -	   5        $ ClientCompanyTextBox��5D      $��4��60   $��1��6=   $	��200��6M   $��5q�6Y   �� 5F   
   $��4��6   $��2��6   $��Candidate phone:��6-   $��Center=�6H   �� 5G   
   -
   5        $ CandidatePhoneTextBox��5G      $��4��61   $��3��6>   $	��200��6N   $��5q�6Z   �� 5I   
   $��5��6   $��0��6   $!��Additional info (optional):��6-   $��Center=�6S   �� 5J   
   -   5        $ AdditionalInfoTextBox��5J      $��5��61   $��1��6>   $��3��6N   $	��660��6b   $��5q�6n   �� 5L   
   -   5        $��6��5L      $��0��6   $��4��6.   $��Preview��6B   $	��100��6T   $��5q�6`   6k   �� 5M   
   -   5         ��Name$ SaveClientButton��5M      $��6��6+   $��0��68   $��2��6H   $��Save client��6\   $	��100��6r   $��5q�6~   $
��left=�6�   6�   �� 5N   
   -   5        $ LoadClientButton��5N      $��6��6+   $��0��68   $��2��6H   $��Load client��6\   $	��100��6r   $��	119,5,0,5q�6~   $
��Left=�6�   6�   �� 5O   
   -   5        $ SaveCandidateButton��5O      $��6��6.   $��3��6;   $��2��6K   $��Save candidate��6_   $	��100��6x   $��5q�6�   $��center=�6�   6�   �� 5P   
   -   5        $ LoadCandidateButton��5P      $��6��6.   $��3��6;   $��2��6K   $��Load candidate��6_   $	��100��6x   $��5q�6�   $��right=�6�   6�   5R      5S      