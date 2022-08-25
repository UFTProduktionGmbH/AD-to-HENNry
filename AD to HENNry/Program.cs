using AD_to_CSV;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;

internal class Program
{
    public static SortedDictionary<String, Info> user = new SortedDictionary<String, Info>();
    public static String outputFile = @"output.csv";
    public static String header = "Identifikation;Vorname;Nachname;Jobtitel;Abteilung;Unternehmensstandort;Email;Telefonnummer;Mobiltelefonnummer;Nutzername";

    public struct Info
    {
        public String identifikation;
        public String vorname;
        public String nachname;
        public String jobtitel;
        public String abteilung;
        public String unternehmensstandort;
        public String email;
        public String telefonnummer;
        public String mobiltelefonnummer;
        public String nutzername;

        public override string ToString() => $"{identifikation};{vorname};{nachname};{jobtitel};{abteilung};{unternehmensstandort};{email};{telefonnummer};{mobiltelefonnummer};{nutzername}";
    }

    static void Main(string[] args)
    {
        //INI erstellen

        bool FileNew = false;

        if (!File.Exists(@"Configuration.ini"))
        {
            var test = File.Create(@"Configuration.ini");
            test.Close();
            FileNew = true;
        }

        IniFile INI = new IniFile("Configuration.ini");

        if (FileNew)
        {
            INI.Write("Domain", "wackershauser.de", "General");
            INI.Write("OUs", "OU=User-Accounts,OU=Benutzer,OU=Werk1,OU=UFT,DC=wackershauser,DC=de;OU=User-Accounts,OU=Benutzer,OU=Werk2,OU=UFT,DC=wackershauser,DC=de", "General");
            INI.Write("DoNotExportMails", "j.doe@uft.de;montage.mobil@uft.de;praktikant@uft.de;schichtfuehrer@uft.de;sortierung@uft.de;versand@uft.de", "General");
            INI.Write("Identifikation", "mail", "Mapping aus AD");
            INI.Write("Vorname", "givenName", "Mapping aus AD");
            INI.Write("Nachname", "sn", "Mapping aus AD");
            INI.Write("Jobtitel", "title", "Mapping aus AD");
            INI.Write("Abteilung", "department", "Mapping aus AD");
            INI.Write("Unternehmensstandort", "l", "Mapping aus AD");
            INI.Write("Email", "mail", "Mapping aus AD");
            INI.Write("Telefonnummer", "ipPhone", "Mapping aus AD");
            INI.Write("Mobiltelefonnummer", "mobile", "Mapping aus AD");
            INI.Write("Nutzername", "sAMAccountName", "Mapping aus AD");
            INI.Write("Token", "NjJkN2NjMWU0ZDJkNTAxMzYwM2MwMGY3OilSeFo2XWRsNW1PJHZEN3lrMTZ3cHgsLHBXLl9bMXVpOzBfQ2Q4UUoqX3NYWiEza1puaF9jZlZSODRpSmVNSVQ=", "Upload");
            INI.Write("Mapping", "externalID,profile-field:firstName,profile-field:lastName,profile-field:jobtitel,profile-field:department,profile-field:standort,eMail,profile-field:telefonnummer,profile-field:mobiltelefonnummer,userName", "Upload");
            INI.Write("ImportTag", "uft", "Upload");
            INI.Write("Dry", "true", "Upload");

            Console.WriteLine("Configuration.ini wurde erstellt und muss angepasst werden!\nBei mehreren Angaben mit ; abtrennen.");

            Console.ReadKey();
            Environment.Exit(0);
        }


        //INI lesen

        string ConfigDomain = "";
        string ConfigOUs = "";
        string ConfigDoNotExportMails = "";

        string ConfigIdentifikation = "";
        string ConfigVorname = "";
        string ConfigNachname = "";
        string ConfigJobtitel = "";
        string ConfigAbteilung = "";
        string ConfigUnternehmensstandort = "";
        string ConfigEmail = "";
        string ConfigTelefonnummer = "";
        string ConfigMobiltelefonnummer = "";
        string ConfigNutzername = "";

        string ConfigToken = "";
        string ConfigMapping = "";
        string ConfigImportTag = "";
        string ConfigDry = "";


        try
        {
            if (INI.KeyExists("Domain", "General"))
            {
                ConfigDomain = INI.Read("Domain", "General");
            }
            else
                throw new Exception("Domain");


            if (INI.KeyExists("OUs", "General"))
            {
                ConfigOUs = INI.Read("OUs", "General");
            }
            else
                throw new Exception("OUs");


            if (INI.KeyExists("DoNotExportMails", "General"))
            {
                ConfigDoNotExportMails = INI.Read("DoNotExportMails", "General");
            }
            else
                throw new Exception("DoNotExportMails");




            if (INI.KeyExists("Identifikation", "Mapping aus AD"))
            {
                ConfigIdentifikation = INI.Read("Identifikation", "Mapping aus AD");
            }
            else
                throw new Exception("Identifikation");


            if (INI.KeyExists("Vorname", "Mapping aus AD"))
            {
                ConfigVorname = INI.Read("Vorname", "Mapping aus AD");
            }
            else
                throw new Exception("Vorname");


            if (INI.KeyExists("Nachname", "Mapping aus AD"))
            {
                ConfigNachname = INI.Read("Nachname", "Mapping aus AD");
            }
            else
                throw new Exception("Nachname"); 


            if (INI.KeyExists("Jobtitel", "Mapping aus AD"))
            {
                ConfigJobtitel = INI.Read("Jobtitel", "Mapping aus AD");
            }
            else
                throw new Exception("Jobtitel");


            if (INI.KeyExists("Abteilung", "Mapping aus AD"))
            {
                ConfigAbteilung = INI.Read("Abteilung", "Mapping aus AD");
            }
            else
                throw new Exception("Abteilung");


            if (INI.KeyExists("Unternehmensstandort", "Mapping aus AD"))
            {
                ConfigUnternehmensstandort = INI.Read("Unternehmensstandort", "Mapping aus AD");
            }
            else
                throw new Exception("Unternehmensstandort");


            if (INI.KeyExists("Email", "Mapping aus AD"))
            {
                ConfigEmail = INI.Read("Email", "Mapping aus AD");
            }
            else
                throw new Exception("Email");


            if (INI.KeyExists("Telefonnummer", "Mapping aus AD"))
            {
                ConfigTelefonnummer = INI.Read("Telefonnummer", "Mapping aus AD");
            }
            else
                throw new Exception("Telefonnummer");


            if (INI.KeyExists("Mobiltelefonnummer", "Mapping aus AD"))
            {
                ConfigMobiltelefonnummer = INI.Read("Mobiltelefonnummer", "Mapping aus AD");
            }
            else
                throw new Exception("Mobiltelefonnummer");


            if (INI.KeyExists("Nutzername", "Mapping aus AD"))
            {
                ConfigNutzername = INI.Read("Nutzername", "Mapping aus AD");
            }
            else
                throw new Exception("Nutzername");



            if (INI.KeyExists("Token", "Upload"))
            {
                ConfigToken = INI.Read("Token", "Upload");
            }
            else
                throw new Exception("Token");


            if (INI.KeyExists("Mapping", "Upload"))
            {
                ConfigMapping = INI.Read("Mapping", "Upload");
            }
            else
                throw new Exception("Mapping");


            if (INI.KeyExists("ImportTag", "Upload"))
            {
                ConfigImportTag = INI.Read("ImportTag", "Upload");
            }
            else
                throw new Exception("ImportTag");


            if (INI.KeyExists("Dry", "Upload"))
            {
                if (INI.Read("Dry", "Upload") == "false" || INI.Read("Dry", "Upload") == "true")
                    ConfigDry = INI.Read("Dry", "Upload");
                else
                    ConfigDry = "true";
            }
            else
                throw new Exception("Dry");
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message + " fehlt in Configuration.ini!");
            Console.ReadKey();
            Environment.Exit(0);
        }
        


        //alt
        try
        {
            List<PrincipalContext> lContext = new List<PrincipalContext>();

            foreach (string item in ConfigOUs.Split(';'))
            {
                lContext.Add(new PrincipalContext(ContextType.Domain, ConfigDomain, item));
            }

            foreach (var context in lContext)
            {

                UserPrincipal userPrincipal = new UserPrincipal(context);
                userPrincipal.Enabled = true;

                PrincipalSearcher src = new PrincipalSearcher(userPrincipal);

                foreach (UserPrincipal found in src.FindAll())
                {
                    DirectoryEntry entry = found.GetUnderlyingObject() as DirectoryEntry;

                    Info userInfo = new Info();

                    if (entry.Properties[ConfigIdentifikation].Value != null)
                        userInfo.identifikation = entry.Properties[ConfigIdentifikation].Value.ToString();

                    if (entry.Properties[ConfigVorname].Value != null)
                        userInfo.vorname = entry.Properties[ConfigVorname].Value.ToString();

                    if (entry.Properties[ConfigNachname].Value != null)
                        userInfo.nachname = entry.Properties[ConfigNachname].Value.ToString();

                    if (entry.Properties[ConfigJobtitel].Value != null)
                        userInfo.jobtitel = entry.Properties[ConfigJobtitel].Value.ToString();

                    if (entry.Properties[ConfigAbteilung].Value != null)
                        userInfo.abteilung = entry.Properties[ConfigAbteilung].Value.ToString();

                    if (entry.Properties[ConfigUnternehmensstandort].Value != null)
                        userInfo.unternehmensstandort = entry.Properties[ConfigUnternehmensstandort].Value.ToString();

                    if (entry.Properties[ConfigEmail].Value != null)
                        userInfo.email = entry.Properties[ConfigEmail].Value.ToString();

                    if (entry.Properties[ConfigTelefonnummer].Value != null)
                        userInfo.telefonnummer = entry.Properties[ConfigTelefonnummer].Value.ToString();

                    if (entry.Properties[ConfigMobiltelefonnummer].Value != null)
                        userInfo.mobiltelefonnummer = entry.Properties[ConfigMobiltelefonnummer].Value.ToString();
                    
                    if (entry.Properties[ConfigNutzername].Value != null)
                        userInfo.nutzername = entry.Properties[ConfigNutzername].Value.ToString();

                    user.Add(found.SamAccountName, userInfo);
                }
            }

            //Output

            var output = new StringBuilder();

            output.AppendLine(header);
            
            foreach (var item in user)
            {
                if (item.Value.email != null && !ConfigDoNotExportMails.Split(';').Contains(item.Value.email))
                {
                    output.AppendLine(item.Value.ToString());
                }
            }

            File.WriteAllText(outputFile, output.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        
        Upload(ConfigToken, ConfigImportTag);
        Thread.Sleep(5000);
        Update(ConfigToken, ConfigMapping, ConfigDry);
        Thread.Sleep(5000);
    }

    public static async Task Upload(string ConfigToken, string ConfigImportTag)
    {
        var handler = new HttpClientHandler();

        handler.AutomaticDecompression = ~DecompressionMethods.All;

        using (var httpClient = new HttpClient(handler))
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), "https://hennry.henn-group.com/api/users/import/csv/upload"))
            {
                request.Headers.TryAddWithoutValidation("Authorization", $"Basic {ConfigToken}");

                var multipartContent = new MultipartFormDataContent();
                multipartContent.Add(new ByteArrayContent(File.ReadAllBytes("./output.csv")), "csv", Path.GetFileName("./output.csv"));
                multipartContent.Add(new StringContent("utf-8"), "encoding");
                multipartContent.Add(new StringContent(";"), "fieldSeparator");
                multipartContent.Add(new StringContent($"csv_import:{ConfigImportTag}"), "partialImportTag");
                request.Content = multipartContent;

                var response = await httpClient.SendAsync(request);
                var resault = await response.Content.ReadAsStringAsync();
                File.WriteAllText("./upload.json", JsonConvert.SerializeObject(JsonConvert.DeserializeObject(resault), Formatting.Indented));
            }
        }
    }

    public static async Task Update(string ConfigToken, string ConfigMapping, string ConfigDry)
    {
        var handler = new HttpClientHandler();

        handler.AutomaticDecompression = ~DecompressionMethods.All;

        using (var httpClient = new HttpClient(handler))
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), "https://hennry.henn-group.com/api/users/import/csv/update"))
            {
                request.Headers.TryAddWithoutValidation("Authorization", $"Basic {ConfigToken}");

                var contentList = new List<string>();
                contentList.Add($"mappings={ConfigMapping}");
                contentList.Add("sendMailsNew=true");
                contentList.Add("sendMailsPending=false");
                contentList.Add("generateRecoveryCodes=false");
                contentList.Add($"dry={ConfigDry}");
                request.Content = new StringContent(string.Join("&", contentList));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/x-www-form-urlencoded");

                var response = await httpClient.SendAsync(request);
                var resault = await response.Content.ReadAsStringAsync();

                File.WriteAllText("./update.json", JsonConvert.SerializeObject(JsonConvert.DeserializeObject(resault), Formatting.Indented));
            }
        }
    }
}