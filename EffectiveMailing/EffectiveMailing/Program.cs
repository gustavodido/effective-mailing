        static StreamWriter fs;

        static void Main(string[] args)
        {
            try
            {
                fs = new StreamWriter(Properties.Settings.Default.Log, true);

                int currAction = 1;
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                CreateTable();
                WriteLog("Criando tabela " + Properties.Settings.Default.TableName);

                WriteLog("Criando conexao com banco");
                OracleConnection conn = new OracleConnection(Properties.Settings.Default.ConnectionString);
                conn.Open();

                WriteLog("Carregando planilhas XLSX");
                foreach (string filename in Directory.GetFiles(Properties.Settings.Default.Directory, "*.xlsx"))
                {
                    if (filename.IndexOf('~') > -1)
                        continue;

                    WriteLog("Carregando arquivo: " + filename);
                    Workbook wb = xlApp.Workbooks.Open(filename);

                    foreach (Worksheet ws in wb.Sheets)
                    {
                        foreach (Range r in ws.get_Range("A2:A20000"))
                            try
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(r.Value)))
                                {
                                    if (ws.get_Range("B1").Value == "FLAG" && !string.IsNullOrEmpty(ws.get_Range("B" + r.Row).Value))
                                        InserRecord(conn, Path.GetFileName(filename), currAction, "", ws.Name, r.Value);
                                    else if (ws.get_Range("C1").Value == "FLAG" && !string.IsNullOrEmpty(ws.get_Range("C" + r.Row).Value))
                                        InserRecord(conn, Path.GetFileName(filename),currAction, r.Value, ws.Name, ws.get_Range("B" + r.Row).Value);
                                    else if (string.IsNullOrEmpty(ws.get_Range("B1").Value))
                                        InserRecord(conn, Path.GetFileName(filename),currAction, "", ws.Name, Convert.ToString(r.Value));
                                }
                                else
                                    break;
                            }
                            catch(Exception ex)
                            {
                                WriteLog(ex.Message + " - " + r.Value);
                            }
                    }

                    wb.Close();
                    currAction++;

                    File.Move(filename, Properties.Settings.Default.MoveDirectory + "\\" + Path.GetFileName(filename));
                    WriteLog("Arquivo movido " + filename);
                }
            
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

            fs.Close();
        }

        static void WriteLog(string log)
        {
            fs.WriteLine(String.Format("[{0}] -> {1}", DateTime.Now.ToString("dd/MM/yyyy hh:mm"), log));
        }

        static void CreateTable()
        {
            using (OracleConnection conn = new OracleConnection(Properties.Settings.Default.ConnectionString))
            {
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = "CREATE TABLE " + Properties.Settings.Default.TableName + " (DT_CARGA DATE, ARQUIVO VARCHAR2(255), ACAO NUMBER, OPERADORA VARCHAR2(255), ATRIBUTO VARCHAR2(100), VALOR VARCHAR(255)) NOLOGGING";
                cmd.CommandType = CommandType.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch(OracleException ex)
                {
                    WriteLog(ex.Message);
                    if (ex.Number != 955)
                        throw ex;
                }               
           }
        }

        static void InserRecord(OracleConnection conn, string arquivo, int acao, string operadora, string atributo, string valor)
        {
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO " + Properties.Settings.Default.TableName + " (DT_CARGA, ARQUIVO, ACAO, OPERADORA, ATRIBUTO, VALOR) VALUES (sysdate, :1, :2, :3, :4, :5)";
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new OracleParameter(":1", OracleDbType.Varchar2)).Value = arquivo;
                cmd.Parameters.Add(new OracleParameter(":2", OracleDbType.Varchar2)).Value = acao;
                cmd.Parameters.Add(new OracleParameter(":3", OracleDbType.Varchar2)).Value = operadora;
                cmd.Parameters.Add(new OracleParameter(":4", OracleDbType.Varchar2)).Value = atributo;
                cmd.Parameters.Add(new OracleParameter(":5", OracleDbType.Varchar2)).Value = valor;
                cmd.ExecuteNonQuery();
        }
