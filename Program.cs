using IronXL;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;

namespace FileGenerator {
    class Program {
        static void Main ( string[] args )
        {
            //string programPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\FILES";
            //string inputsPath = programPath + @"\INPUTS";
            //string archive = programPath + @"\ARCHIVES";
            //string outputsPath = programPath + @"\OUTPUTS";
            //string sharePath = @"\\192.168.20.19\everyone\Settlement\THIRD PARTY TRANSFER";
            //MoveExcelFiles(sharePath, inputsPath);

            //DirectoryInfo pathInfo = new DirectoryInfo(inputsPath);
            //var excelFiles = pathInfo.GetFiles();

            //if (excelFiles.ToList().Count > 0 && excelFiles.First().Extension != ".db")
            //{
            //    excelFiles.ToList().ForEach(excel =>
            //    {
            //        if (excel.Extension == "xls" || excel.Extension == "xlsx")
            //        {
            //            DataTable table = ReadExcel(inputsPath + $@"\{excel}");
            //            string details = excel.Name.Replace(".xlsx", string.Empty);
            //            string account = details.Split('_')[0];
            //            string customerName = details.Split('_')[1];

            //            GenerateWBCFormat(table, account, customerName, outputsPath);
            //            GenerateBASISFormat(table, account, customerName, outputsPath);

            //            Console.WriteLine("MOVE GENERATED FILES TO SETTLEMENT SHARE");
            //            string dateTime = DateTime.Now.ToString("yyyyMMdd");

            //            File.Move(outputsPath + $@"\Virement.SALAIRES.{dateTime}.010.LOT", sharePath + $@"\OUTPUT\Virement.SALAIRES.{dateTime}.010.LOT");
            //            File.Move(outputsPath + $@"\SALAIRES_{customerName}_{dateTime}.txt", sharePath + $@"\OUTPUT\SALAIRES_{customerName}_{dateTime}.txt");

            //            Console.WriteLine(@"FILES HAVE BEEN MOVED SUCCESSFULLY CHECK: \\192.168.20.19\everyone\Settlement\THIRD PARTY TRANSFER\OUTPUT");

            //            File.Move(inputsPath + $@"\{excel}", archive + $@"\{excel}");
            //        }
            //    });
            //    Console.WriteLine("FILES GENERATED!\n-----------------------------------------------------");
            //}
            //else
            //{
            //    Console.WriteLine("NO FILES IN INPUTS DIRECTORY\n-----------------------------------------------------");
            //}
            using (Timer timer = new Timer(60000))  //  (1000 * 120 ) for 2 minutes
            {
                //Add event
                timer.Elapsed += (sender, e) => ProcessFileGeneration(sender, e);

                timer.Start();

                Console.WriteLine("--------------------------------------Timer is started--------------------------------------");

                // Type <<ENTER>> if you want to stop the process
                Console.ReadLine();
            }

        }

        public static void ProcessFileGeneration( Object source, ElapsedEventArgs e )
        {
            string programPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\FILES" ;
            string inputsPath = programPath + @"\INPUTS";
            string archive = programPath + @"\ARCHIVES";
            string outputsPath = programPath + @"\OUTPUTS";
            string sharePath = @"\\192.168.20.19\everyone\Settlement\THIRD PARTY TRANSFER";
            MoveExcelFiles( sharePath, inputsPath );

            DirectoryInfo pathInfo = new DirectoryInfo(inputsPath);
            var excelFiles = pathInfo.GetFiles();

            if ( excelFiles.ToList().Count > 0 && excelFiles.First().Extension != ".db" )
            {
                excelFiles.ToList().ForEach( excel => {
                    if(excel.Extension == ".xls" || excel.Extension == ".xlsx" )
                    {
                        DataTable table = ReadExcel(inputsPath + $@"\{excel}");
                        string details = excel.Name.Replace(".xlsx", string.Empty);
                        string account = details.Split('_')[0];
                        string customerName = details.Split('_')[1];

                        GenerateWBCFormat( table, account, customerName, outputsPath );
                        GenerateBASISFormat( table, account, customerName, outputsPath );

                        Console.WriteLine( "MOVE GENERATED FILES TO SETTLEMENT SHARE" );
                        string dateTime = DateTime.Now.ToString("yyyyMMdd");

                        File.Move( outputsPath + $@"\Virement.TRANSFERTS.{dateTime}.010.LOT", sharePath + $@"\OUTPUT\Virement.TRANSFERTS.{dateTime}.010.LOT" );
                        File.Move( outputsPath + $@"\TRANSFERTS_{customerName}_{dateTime}.DAT", sharePath + $@"\OUTPUT\TRANSFERTS_{customerName}_{dateTime}.DAT" );

                        Console.WriteLine( @"FILES HAVE BEEN MOVED SUCCESSFULLY CHECK: \\192.168.20.19\everyone\Settlement\THIRD PARTY TRANSFER\OUTPUT" );

                        File.Move( inputsPath + $@"\{excel}", archive + $@"\{excel}" );
                    }
                } );
                Console.WriteLine( "FILES GENERATED!\n-----------------------------------------------------" );
            }
            else
            {
                Console.WriteLine( "NO FILES IN INPUTS DIRECTORY\n-----------------------------------------------------" );
            }
        }

        public static void MoveExcelFiles (string source, string destination)
        {
            Console.WriteLine( "MOVE EXCEL FILES" );
            DirectoryInfo sourcePathInfo = new DirectoryInfo(source + @"\INPUT");

            Console.WriteLine( "CHECK FOR FILES..." );

            var excelFiles = sourcePathInfo.GetFiles();
            if(excelFiles.Count() != 0 )
            {
                excelFiles.ToList().ForEach( excel => File.Move( source + $@"\INPUT\{excel}", destination + $@"\{excel}" ) );
            }
            else
            {
                Console.WriteLine( "NO FILES YET" );
            }
        }

        private static DataTable ReadExcel(string fileName )
        {
            WorkBook workbook = WorkBook.Load(fileName);

            WorkSheet sheet = workbook.DefaultWorkSheet;

            return sheet.ToDataTable( true );
        }

        private static void GenerateWBCFormat (DataTable data, string accountToDebit, string customerName, string outputPath)
        {
            Console.WriteLine( "GENERATE WEBCLEARING FILE" );

            for ( int i = data.Rows.Count - 1; i >= 0; i-- )
            {
                DataRow item = data.Rows[i];
                if ( item[0].ToString() == string.Empty || item[0].ToString().ToUpper() == "TOTAL" )
                    item.Delete();
            }
            data.AcceptChanges();

            List<string> transactionsList = new List<string>();
            
            string bankDate = string.Empty;
            OracleDataReader oraReader = GetBasisDate(999);
            if ( oraReader.HasRows )
            {
                oraReader.Read();
                bankDate = oraReader.GetString( 0 );
            }

            string dateTime = DateTime.Now.ToString("yyyyMMdd");

            string nuban = accountToDebit;
            string custName = RemoveDiacritics(Regex.Replace(customerName, @"[^\w\.@ ]", string.Empty));
            string address = "ABIDJAN";
            string beneficiaryAcct = string.Empty;
            string beneficiaryName = string.Empty;
            string transactionAmt = string.Empty;
            string remarks = string.Empty;
            string newLine = string.Empty;
            string space =  string.Empty;

            foreach (DataRow item in data.Rows )
            {
                //string entry = item[0].ToString();

                if (item[0].ToString() != string.Empty || item[0].ToString().ToUpper() != "TOTAL" )
                {
                    beneficiaryAcct = item[3].ToString().Replace( " ", string.Empty );
                    beneficiaryName = RemoveDiacritics( Regex.Replace( item[1].ToString(), @"[^\w\.@ ]", string.Empty ) );
                    transactionAmt = item[4].ToString().Replace( ",", string.Empty );
                    transactionAmt = transactionAmt.Replace( ".", string.Empty );
                    remarks = "VIREMENT DO " + custName + " FAV " + beneficiaryName;

                    newLine = "010" + space.PadRight( 20, ' ' ) + nuban + custName.PadRight( 50, ' ' ) + space.PadRight( 70, ' ' ) + space.PadRight( 3, ' ' ) +
                        bankDate + space.PadRight( 11, ' ' ) + beneficiaryAcct + beneficiaryName.PadRight( 50, ' ' ) + address.PadRight( 70, ' ' ) +
                        transactionAmt.PadLeft( 13, '0' ) + space.PadRight( 2, ' ' ) + remarks.PadRight( 70, ' ' );

                    transactionsList.Add( newLine );
                }
            }

            string writingPath = outputPath + $@"\Virement.TRANSFERTS.{dateTime}.010.LOT";
            TextWriter tw = new StreamWriter(writingPath);
            foreach(string s in transactionsList )
            {
                if(s != transactionsList.Last() )
                {
                    tw.WriteLine( s );
                }
                else
                {
                    tw.Write( s );
                }
            }

            tw.Close();
        }

        private static void GenerateBASISFormat( DataTable data, string accountToDebit, string customerName, string outputPath )
        {
            Console.WriteLine( "GENERATE BASIS FILE" );

            for ( int i = data.Rows.Count - 1; i >= 0; i-- )
            {
                DataRow item = data.Rows[i];
                if ( item[0].ToString() == string.Empty || item[0].ToString().ToUpper() == "TOTAL" )
                    item.Delete();
            }
            data.AcceptChanges();

            string acctDetails = GetCustDetails(accountToDebit);
            var acctDetailsArr = acctDetails.Split('_');

            string branchCode = acctDetailsArr[0];
            string customerNum = acctDetailsArr[1];
            string curCode = acctDetailsArr[2];
            string ledgerCode = acctDetailsArr[3];
            string subAcctCode = acctDetailsArr[4];
            string beneficiaryName = string.Empty;
            string amount = string.Empty;
            string remarks = string.Empty;
            string checkDigit = GetCheckDigit(branchCode, customerNum, curCode, ledgerCode).ToString();

            DataView dv = data.DefaultView;
            dv.Sort = "NOM DE LA BANQUE ASC";
            DataTable data2 = dv.ToTable();

            List<string> accountNumbers = data2.AsEnumerable().Select(d => d.Field<string>("NUMERO DE COMPTE")).ToList();
            List<string> banksCode = new List<string>();

            foreach(var val in accountNumbers )
            {
                if(!string.IsNullOrEmpty(val))
                    banksCode.Add( val.Substring( 0, 5 ) );
            }

            banksCode = banksCode.Distinct().ToList();

            List<string> banks = data2.AsEnumerable().Select(d => d.Field<string>("NOM DE LA BANQUE")).Distinct().ToList();
            List<dynamic> amounts = data2.AsEnumerable().Select(d => d[4]).ToList();
            int totalAmount = 0;
            foreach (var salary in amounts )
            {
                string transactionAmt = salary.ToString().Replace( ",", string.Empty );
                transactionAmt = transactionAmt.Replace( ".", string.Empty );

                totalAmount += Convert.ToInt32( transactionAmt );
            }

            List<string> transactionsList = new List<string>();

            string newLine = string.Empty;
            string feesLine1 = string.Empty;
            string feesLine2 = string.Empty;
            string tobLine1 = string.Empty;
            string tobLine2 = string.Empty;

            string account_409 = ConfigurationManager.AppSettings["AcctToCredit"].ToString();
            string checkDigit409 = GetCheckDigit("201", "70", "1", "409").ToString();

            if(data2.Rows.Count <= 1 )
            {
                newLine = branchCode.PadLeft( 4, '0' ) + customerNum.PadLeft( 7, '0' ) + checkDigit + ledgerCode.PadLeft( 4, '0' )
                        + subAcctCode.PadLeft( 3, '0' ) + 1 + totalAmount.ToString().PadLeft( 15, '0' ) + "000000000000000" + "VIREMENT DO " + customerName + $" FAV {data.Rows[0][1]}";
            }
            else
            {
                newLine = branchCode.PadLeft( 4, '0' ) + customerNum.PadLeft( 7, '0' ) + checkDigit + ledgerCode.PadLeft( 4, '0' )
                        + subAcctCode.PadLeft( 3, '0' ) + 1 + totalAmount.ToString().PadLeft( 15, '0' ) + "000000000000000" + "VIREMENT DO " + customerName + " FAV BENEFICIAIRES DIVERS";
            }
            
            transactionsList.Add( newLine );

            List<VirementInfo> virements = new List<VirementInfo>();
            foreach(string bank in banksCode )
            {
                VirementInfo virement = new VirementInfo{BankCode = bank};
                virement.Items = new List<DataRow>();
                foreach( DataRow item in data2.Rows )
                {
                    if( !( string.IsNullOrEmpty( item[3].ToString() ) ) )
                    {
                        if(bank == item[3].ToString().Substring(0, 5) )
                        {
                            virement.Items.Add( item );
                        }
                    }
                }
                virements.Add( virement );
            }

            foreach(var item in virements )
            {
                List<string> customers = new List<string>();
                foreach (var item2 in item.Items )
                {
                    beneficiaryName = RemoveDiacritics( Regex.Replace( item2[1].ToString(), @"[^\w\.@ ]", string.Empty ) );
                    customers.Add(beneficiaryName);
                    amount = item2[4].ToString().Replace( ",", string.Empty );
                    amount = amount.Replace( ".", string.Empty );
                    remarks = "VIREMENT DO " + customerName + " FAV " + beneficiaryName;
                    

                    newLine = account_409.Split( '/' )[0].PadLeft( 4, '0' ) + account_409.Split( '/' )[1].PadLeft( 7, '0' ) + checkDigit409 + account_409.Split( '/' )[3].PadLeft( 4, '0' )
                        + account_409.Split( '/' )[4].PadLeft( 3, '0' ) + 2 + amount.PadLeft( 15, '0' ) + "000000000000000" + remarks;

                    

                    if (item2 != item.Items.Last() )
                    {
                        transactionsList.Add( newLine );
                    }
                    else
                    {
                        transactionsList.Add( newLine );
                        if ( item.BankCode != "CI163" )
                        {
                            feesLine1 = branchCode.PadLeft( 4, '0' ) + customerNum.PadLeft( 7, '0' ) + checkDigit + ledgerCode.PadLeft( 4, '0' )
                            + subAcctCode.PadLeft( 3, '0' ) + 1 + "5000".PadLeft( 15, '0' ) + "000000000000000" + "FRAIS VIREMENT DO " + customerName + $" FAV {item2[2]}";

                            feesLine2 = "201".PadLeft( 4, '0' ) + "0".PadLeft( 7, '0' ) + GetCheckDigit( "201", "0", "1", "8701" ) + "8701".PadLeft( 4, '0' )
                                + "0".PadLeft( 3, '0' ) + 2 + "5000".PadLeft( 15, '0' ) + "000000000000000" + "FRAIS VIREMENT DO " + customerName + $" FAV {item2[2]}";

                            tobLine1 = branchCode.PadLeft( 4, '0' ) + customerNum.PadLeft( 7, '0' ) + checkDigit + ledgerCode.PadLeft( 4, '0' )
                                + subAcctCode.PadLeft( 3, '0' ) + 1 + "500".PadLeft( 15, '0' ) + "000000000000000" + "TOB VIREMENT DO " + customerName + $" FAV {item2[2]}";

                            tobLine2 = "201".PadLeft( 4, '0' ) + "0".PadLeft( 7, '0' ) + GetCheckDigit( "201", "0", "1", "4522" ) + "4522".PadLeft( 4, '0' )
                                + "0".PadLeft( 3, '0' ) + 2 + "500".PadLeft( 15, '0' ) + "000000000000000" + "TOB VIREMENT DO " + customerName + $" FAV {item2[2]}";
                            transactionsList.Add( feesLine1 );
                            transactionsList.Add( feesLine2 );
                            transactionsList.Add( tobLine1 );
                            transactionsList.Add( tobLine2 );
                        }
                    }
                }
            }

            string dateTime = DateTime.Now.ToString("yyyyMMdd");

            string writingPath = outputPath + $@"\TRANSFERTS_{customerName}_{dateTime}.DAT";
            TextWriter tw = new StreamWriter(writingPath);
            int lineNb = transactionsList.Count();
            int index = 0;
            foreach ( string s in transactionsList )
            {
                string test = transactionsList.Last();
                if ( index != lineNb - 1 )
                {
                    tw.WriteLine( s );
                }
                else
                {
                    tw.Write( s );
                }

                index++;
            }

            tw.Close();
        }

        private static string RemoveDiacritics ( string text )
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder(capacity: normalizedString.Length);

            for ( int i = 0; i < normalizedString.Length; i++ )
            {
                char c = normalizedString[i];
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if ( unicodeCategory != UnicodeCategory.NonSpacingMark )
                {
                    stringBuilder.Append( c );
                }
            }

            return stringBuilder
                .ToString()
                .Normalize( NormalizationForm.FormC );
        }

        public static OracleDataReader GetBasisDate ( int branch_code )
        {
            OracleDataReader oraReader;
            OracleConnection Conn = new OracleConnection(ConfigurationManager.AppSettings["OracleConString"].ToString());
            string query = "Select to_char(bank_date,'yyyymmdd') from process where bra_code= " + branch_code;

            OracleCommand comm = new OracleCommand(query, Conn);
            comm.CommandType = CommandType.Text;

            //open connection
            if ( Conn.State != ConnectionState.Open )
            {
                Conn.Open();
            }
            oraReader = comm.ExecuteReader();

            return oraReader;
        }

        private static string GetCustDetails(string nuban )
        {
            string response = string.Empty;
            OracleDataReader oraReader;
            OracleConnection Conn = new OracleConnection(ConfigurationManager.AppSettings["OracleConString"].ToString());
            string query = $"SELECT bra_code, cus_num, cur_code, led_code, sub_acct_code FROM map_acct WHERE map_acc_no = '{nuban}'";

            OracleCommand comm = new OracleCommand(query, Conn);
            comm.CommandType = CommandType.Text;
            Conn.Open();
            oraReader = comm.ExecuteReader();
            while ( oraReader.Read() )
            {
                response = $"{oraReader["BRA_CODE"]}_{oraReader["CUS_NUM"]}_" +
                    $"{oraReader["CUR_CODE"]}_{oraReader["LED_CODE"]}_{oraReader["SUB_ACCT_CODE"]}";
            }
            oraReader.Close();
            Conn.Close();
            return response;
        }

        private static int GetCheckDigit (string branchCode, string customerNum , string currencyCode, string ledgerCode )
        {
            string connStringODB = ConfigurationManager.AppSettings["OracleConString"].ToString();
            string query = $"select che_dig from account where bra_code = {branchCode} and cus_num = {customerNum} and cur_code = {currencyCode} and led_code = {ledgerCode}";

            OracleConnection oracleConnection = new OracleConnection(connStringODB);

            oracleConnection.Open();
            OracleCommand oracleCmd = new OracleCommand(query, oracleConnection);

            OracleDataReader dr = oracleCmd.ExecuteReader();
            int chequeDigit = 0;
            while ( dr.Read() )
            {
                chequeDigit = Convert.ToInt32( dr["CHE_DIG"].ToString() );
            }
            dr.Close();
            oracleConnection.Close();
            return chequeDigit;
        }
    }
}
