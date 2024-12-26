English

# AD-UserManagementTool

## Overview
AD-UserManagementTool is a comprehensive PowerShell tool designed to manage and report on user accounts in an Active Directory (AD) environment. 
This tool provides AD administrators with powerful capabilities to view, filter, edit, and report on user accounts.

## Features
1. Remote Connection
2. ![1](https://github.com/user-attachments/assets/fe528dfa-4ae5-4036-b924-ecb10749bc52)

3. User List Display
4. ![2](https://github.com/user-attachments/assets/402f3c8f-5817-4d52-b9d2-cec78a6ba5d2)

5. Advanced Filtering
6. ![3](https://github.com/user-attachments/assets/13926314-585f-478a-9457-feb9ddb90fa5)

7. User Details
8. ![7](https://github.com/user-attachments/assets/fd7c7804-f4e2-46eb-a995-1db2d8d3db21)

9. Password Management
10. ![Screenshot 2024-12-26 204653](https://github.com/user-attachments/assets/01bc5427-3563-4f20-8518-572b0664f254)

11. Account Locking/Unlocking
12. ![Screenshot 2024-12-26 204807](https://github.com/user-attachments/assets/3946341b-450a-495e-8353-055c32389d17)

13. Data Export (CSV format)
14. ![Screenshot 2024-12-26 204853](https://github.com/user-attachments/assets/0fbd3b88-eda3-4ce0-a98a-2aea5db05d45)
![5](https://github.com/user-attachments/assets/7615d372-36d3-43b6-a914-7fa43c6644d4)

15. Real-Time Statistics
16. Security Measures and Operation Logging

## System Requirements
- Windows PowerShell 5.1 or higher
- Active Directory PowerShell module
- Appropriate permissions to connect to domain controller or AD

## Installation and Launch
1. Download the script and save it to a secure location.
2. Open PowerShell as an administrator.
3. Run the script: `.\AD-UserManagementTool.ps1`

## Usage Instructions

### 1. Remote Connection
- When the script starts, you will be prompted to connect to a domain controller.
- Enter the server name (IP or DNS), username, and password.
- If the connection is successful, the main interface will open.

### 2. Main Panel
The main panel provides an overview of AD users:
- Total Users
- Active Users
- Deactivated Accounts
- Expired Passwords
- Accounts with No Password Set
- Valid Passwords

### 3. User List
- All AD users are displayed in a data grid.
- Columns: Username, Email, Department, Last Login, Status, Password Status, OU, Groups

### 4. Filtering and Search
- You can perform quick searches using the search box at the top.
- Use the expandable filter panel for detailed filtering:
  - Status (Active/Inactive)
  - Password Status
  - Last Login Date

### 5. User Operations
Right-click on a user to perform the following operations:
- View User Details
- Reset Password
- Lock/Unlock Account

### 6. Data Export
- Use the "Export" button to export the displayed user list in CSV format.

### 7. Refresh
- Use the "Refresh" button to update AD data.

## Security Notes
- This tool has powerful capabilities. It should only be used by authorized administrators.
- All operations are logged in the `C:\AD-Reports-logs` folder.
- Ensure you comply with your organization's security policies when using the script.

## Tips
- Regularly use the "Refresh" function before using the tool.
- In large AD environments, optimize your searches by using the filtering features.
- Always carefully read confirmation dialogs before making significant changes.

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Azərbaycan dilində

# AD-İstifadəçiİdarəetməAləti

## Ümumi Baxış
AD-İstifadəçiİdarəetməAləti, Active Directory (AD) mühitində istifadəçi hesablarını idarə etmək və hesabat vermək üçün hazırlanmış geniş bir 
PowerShell alətidir. Bu alət, AD administratorlarına istifadəçi hesablarını göstərmək, filtərləmək, redaktə etmək və hesabat vermək üçün güclü imkanlar təqdim edir.

## Xüsusiyyətlər
1. Uzaqdan Qoşulma
2. İstifadəçi Siyahısını Göstərmə
3. Təkmilləşdirilmiş Filtrasiya
4. İstifadəçi Təfərrüatları
5. Şifrə İdarəetməsi
6. Hesab Kilidləmə/Kilidi Açma
7. Məlumatların İxracı (CSV formatında)
8. Real Vaxt Statistikası
9. Təhlükəsizlik Tədbirləri və Əməliyyat Jurnalı

## Sistem Tələbləri
- Windows PowerShell 5.1 və ya daha yüksək versiya
- Active Directory PowerShell modulu
- Domain controller-ə və ya AD-yə qoşulmaq üçün müvafiq icazələr

## Quraşdırma və İşə Salma
1. Skripti yükləyin və etibarlı bir yerə qeyd edin.
2. PowerShell-i administrator kimi açın.
3. Skripti işə salın: .\AD-İstifadəçiİdarəetməAləti.ps1

## İstifadə Təlimatı

### 1. Uzaqdan Qoşulma
- Skript başladıqda, bir domain controller-ə qoşulmağınız tələb olunacaq.
- Server adını (IP və ya DNS), istifadəçi adını və şifrəni daxil edin.
- Əgər qoşulma uğurlu olarsa, əsas interfeys açılacaq.

### 2. Əsas Panel
Əsas panel AD istifadəçiləri haqqında ümumi məlumat təqdim edir:
- Ümumi İstifadəçilər
- Aktiv İstifadəçilər
- Deaktiv Edilmiş Hesablar
- Müddəti Keçmiş Şifrələr
- Şifrə Təyin Edilməmiş Hesablar
- Etibarlı Şifrələr

### 3. İstifadəçi Siyahısı
- Bütün AD istifadəçiləri bir data grid-də göstərilir.
- Sütunlar: İstifadəçi Adı, E-poçt, Şöbə, Son Giriş, Status, Şifrə Statusu, OU, Qruplar

### 4. Filtrasiya və Axtarış
- Yuxarı hissədəki axtarış qutusu ilə sürətli axtarış edə bilərsiniz.
- Genişləndirilə bilən filtr paneli ilə ətraflı filtrasiya edə bilərsiniz:
  - Status (Aktiv/Deaktiv)
  - Şifrə Statusu
  - Son Giriş Tarixi

### 5. İstifadəçi Əməliyyatları
İstifadəçiyə sağ kliklə aşağıdakı əməliyyatları yerinə yetirə bilərsiniz:
- İstifadəçi Təfərrüatlarını Göstərmə
- Şifrəni Sıfırlama
- Hesabı Kilidləmə/Kilidini Açma

### 6. Məlumatların İxracı
- "İxrac" düyməsi ilə göstərilən istifadəçi siyahısını CSV formatında ixrac edə bilərsiniz.

### 7. Yeniləmə
- "Yenilə" düyməsi ilə AD məlumatlarını yeniləyə bilərsiniz.

## Təhlükəsizlik Qeydləri
- Bu alət güclü imkanlara malikdir. Yalnız səlahiyyətli administratorlar tərəfindən istifadə edilməlidir.
- Bütün əməliyyatlar C:\AD-Reports-logs qovluğunda jurnallaşdırılır.
- Skriptdən istifadə edərkən təşkilatınızın təhlükəsizlik siyasətlərinə əməl etdiyinizdən əmin olun.

## Məsləhətlər
- Mütəmadi olaraq alətdən istifadə etməzdən əvvəl "Yenilə" funksiyasından istifadə edin.
- Böyük AD mühitlərində, filtrasiya xüsusiyyətlərindən istifadə edərək axtarışlarınızı optimallaşdırın.




