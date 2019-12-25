using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppTerminal.Convert.XlsToCsv.Nethome
{
    class ReadFileXls
    {
        public async Task<bool> execute(string CONTENT)
        {
            await createJsonFile(CONTENT);
            await compressionJsonFile();
            return true;
        }

        private async Task<bool> createJsonFile(string seriaisContent)
        {
            try
            {
                string seriaisJson = "{\"serias_atlas\":" + createObjest(seriaisContent) + "}";
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "NetHome\\IMPORTACAO_ATLAS_EMPREITEIRA.json", seriaisJson);
                seriaisContent = null;
                seriaisJson = null;
            }
            catch { }
            return true;
        }
        private async Task<bool> compressionJsonFile()
        {
            try
            {
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "zip\\fileAtlas.zip"))
                {
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + "zip\\fileAtlas.zip");
                }

                ZipFile.CreateFromDirectory(AppDomain.CurrentDomain.BaseDirectory + "NetHome", AppDomain.CurrentDomain.BaseDirectory + "zip\\fileAtlas.zip");
            }
            catch { }
            return true;
        }

        private string createObjest(string CONTENT)
        {
            try
            {
                List<AtlasTerminais> terminais = new List<AtlasTerminais>();
                using (HtmlTableToArray tableToArray = new HtmlTableToArray())
                {
                    StringBuilder stringBuilder = new StringBuilder();

                    stringBuilder.Append(CONTENT);
                    stringBuilder.Replace(" class=\"xString\"", "")
                         .Replace(" bgcolor=\"#FFFFFF\"", "")
                         .Replace("<b style=\"color: #FFFFFF\">", "")
                         .Replace("</b>", "")
                         .Replace(" bgcolor=\"#000080\"", "")
                         .Replace(" bgcolor=\"#000080\"", "")
                         .Replace("</TR>", "")
                         .Replace("</TD>", "")
                         .Replace("</TABLE>", "");

                    var trList = tableToArray.trToArray(stringBuilder.ToString());

                    foreach (var tr in trList)
                    {
                        string[] tdsList = tr.Split(new string[] { "<TD>" }, StringSplitOptions.None);
                        tdsList = tdsList.Skip(1).ToArray();
                        AtlasTerminais terminaisAtlas = new AtlasTerminais();
                        terminaisAtlas.Tipo = tdsList[0].Trim();
                        terminaisAtlas.Modelo = tdsList[1].Trim();
                        terminaisAtlas.CodigoItemJDE = tdsList[2].Trim();
                        terminaisAtlas.CodigoMaterialSAP = tdsList[3].Trim();
                        terminaisAtlas.NumerSerie = tdsList[4].Trim();
                        terminaisAtlas.EndPrincipal = tdsList[5].Trim();
                        terminaisAtlas.Operacao = tdsList[6].Trim();
                        terminaisAtlas.NomeDoLocal = tdsList[7].Trim();
                        terminaisAtlas.Perfil = tdsList[8].Trim();
                        terminaisAtlas.CodigoFornecedorJDE = tdsList[9].Trim();
                        terminaisAtlas.CodigoFornecedorSAP = tdsList[10].Trim();
                        terminaisAtlas.Estado = tdsList[11].Trim();
                        terminaisAtlas.DataUltimaAlteracao = tdsList[12].Trim();
                        terminaisAtlas.Responsavel = tdsList[13].Trim();
                        terminaisAtlas.NumeroContrato = tdsList[14].Trim();
                        terminaisAtlas.ClassificacaoMaterial = tdsList[15].Trim();
                        terminaisAtlas.EmpresaMaterial = tdsList[16].Trim();
                        terminais.Add(terminaisAtlas);
                    }

                    trList = null;

                    terminais.RemoveAt(0);
                  
                    string seriais = JsonConvert.SerializeObject(terminais);

                    terminais = null;

                    CONTENT = string.Empty;

                    tableToArray.Dispose();

                    return seriais;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

        }
    }
}
