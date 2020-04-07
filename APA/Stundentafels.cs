using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace APA
{
    public class Stundentafels:List<Stundentafel>
    {
        public Stundentafels()
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConU))
            {
                string queryString = @"SELECT 
PeriodsTable.PERIODS_TABLE_ID, 
PeriodsTable.Name, 
PeriodsTable.Longname, 
PeriodsTable.PerTabElement1
FROM PeriodsTable
WHERE (((PeriodsTable.SCHOOLYEAR_ID)=" + Global.AktSjUnt + ") AND ((PeriodsTable.Deleted)=No)) ORDER BY Name;";

                OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                oleDbConnection.Open();
                OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();
                
                while (oleDbDataReader.Read())
                {
                    Stundentafel stundentafel = new Stundentafel();

                    var lernbereich = "BB";

                    try
                    {
                        stundentafel.IdUntis = oleDbDataReader.GetInt32(0);
                        stundentafel.Name = Global.SafeGetString(oleDbDataReader, 1);                        
                        stundentafel.Langname = Global.SafeGetString(oleDbDataReader, 2);
                        
                        var elemente = (Global.SafeGetString(oleDbDataReader, 3)).Split(',');
                        int i = 1;

                        foreach (var element in elemente)
                        {
                            if (element.Split('~')[2] == "D")
                            {
                                lernbereich = "BÜ";
                            }
                            stundentafel.Fachs.Add(new Fach(i, element.Split('~')[2], lernbereich));
                            i++;
                        }                        
                    }
                    catch (Exception ex)
                    {
                     
                    }

                    if (!(from s in this where s.IdUntis == stundentafel.IdUntis select s).Any())
                    {
                        this.Add(stundentafel);
                    }                    
                };
                Console.WriteLine(("Stundentafeln " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                oleDbDataReader.Close();
                oleDbConnection.Close();
            }
        }
    }
}