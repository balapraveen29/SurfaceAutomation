using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.IO;

namespace SurfaceAutomation
{
    public class clsDBAccess
    {
        public System.Data.DataTable SelectTable(string strQuery, Hashtable hat, string queryType, string conString)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlConnection con = new SqlConnection(conString);
            SqlCommand cmd = new SqlCommand();
            try
            {
                con.Close();
                con.Open();
                SqlDataAdapter ad = new SqlDataAdapter(strQuery, con);
                if(queryType.Trim().ToLower()=="sp")
                {
                    ad.SelectCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    foreach(DictionaryEntry hatVal in hat)
                    {
                        ad.SelectCommand.Parameters.AddWithValue(hatVal.Key.ToString(), hatVal.Value.ToString());
                    }
                }
                else
                {
                    ad.SelectCommand.CommandType = System.Data.CommandType.Text;
                }
                ad.Fill(dt);
            }
            catch
            {

            }
            finally
            {
                con.Close();
            }
            return dt;
        }

        public int WriteTable(string strQuery, Hashtable hat, string queryType, string conString)
        {
            int cnt = 0;
            SqlConnection con = new SqlConnection(conString);
            SqlCommand cmd = new SqlCommand();
            try
            {
                con.Close();
                con.Open();
                cmd = new SqlCommand(strQuery, con);
                if (queryType.Trim().ToLower() == "sp")
                {
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    foreach (DictionaryEntry hatVal in hat)
                    {
                        cmd.Parameters.AddWithValue(hatVal.Key.ToString(), hatVal.Value.ToString());
                    }
                }
                else
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                }
                cnt = cmd.ExecuteNonQuery();
            }
            catch
            {

            }
            finally
            {
                con.Close();
            }
            return cnt;
        }

        public string Encrypt(string strText)
        {
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] clearBytes = Encoding.Unicode.GetBytes(strText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using(MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    strText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return strText;
        }

        public string Decrypt(string strText)
        {
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] cipherBytes = Convert.FromBase64String(strText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    strText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return strText;
        }
    }    
}
  