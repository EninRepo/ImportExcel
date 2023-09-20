# ImportExcel
How to Import Excel file Using OLEDBConnection Nuget Package in Asp.net Core 3.1
OLEDBConnection version 6.0
other depandence is defualt 3.15....
check the platform compatibility x64, x84 to solve the extention
    switch (extenstion)
                    {
                        case ".xls":
                            conString = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=YES'"; // If your os support use condition to check
                            break;
                        case ".xlsx":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                    }
