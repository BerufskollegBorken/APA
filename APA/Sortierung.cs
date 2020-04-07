using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace APA
{
    public class Sortierung : List<Fach>
    {
        public Sortierung()
        {
            object aaa;
            try
            {
                using (OdbcConnection connection = new OdbcConnection(Global.ConAtl))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT DBA.fachkomb_kopf.bezeichnung AS Bezeichnung,
DBA.fachkomb_einzel.position_1 AS Position1,
DBA.fach.kuerzel AS Fach
FROM(DBA.fachkomb_kopf JOIN DBA.fachkomb_einzel ON DBA.fachkomb_kopf.fko_id = DBA.fachkomb_einzel.fko_id) JOIN DBA.fach ON DBA.fachkomb_einzel.fa_id = DBA.fach.fa_id
where fachkomb_kopf.aktiv_jn = 'J'
ORDER BY DBA.fachkomb_kopf.fko_id ASC ,
DBA.fachkomb_einzel.position_1 ASC"
, connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.klasse");
                    
                    foreach (DataRow theRow in dataSet.Tables["DBA.klasse"].Rows)
                    {
                        Fach fach = new Fach();

                        fach.Bezeichnung = theRow["Bezeichnung"] == null ? "" : theRow["Bezeichnung"].ToString();                       
                        fach.FachkürzelAtlantis = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                        fach.Position1 = theRow["Position1"] == null ? -99 : Convert.ToInt32(theRow["Position1"]);
                        this.Add(fach);
                    }
                    connection.Close();
                }                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}