﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace DataAccess
{
    public class ManagerDA
    {
        private string GetConectionString()
        {
            /* Modificacion Alvaro Araujo 07/06/2019 */
            return ConfigurationManager.ConnectionStrings["SqlProvider"].ConnectionString;
        }

        public SqlConnection GetConnection()
        {
            SqlConnection con = new SqlConnection(GetConectionString());
            con.Open();
            return con;
        }

        public SqlParameter[] GetParameters(int cantidad)
        {
            List<SqlParameter> arr = new List<SqlParameter>();
            int contador;
            for (contador = 0; contador < cantidad; contador++)
            {
                arr.Add(new SqlParameter());
            }
            return arr.ToArray();
        }

        private void SetParameters(SqlCommand cmd, SqlParameter[] arr)
        {
            if (arr != null)
            {
                foreach (SqlParameter obj in arr)
                    cmd.Parameters.Add(obj);
            }
        }

        public int ExecuteNonQuery(string query, SqlParameter[] arr, CommandType tipo)
        {
            SqlConnection con = GetConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            cmd.CommandType = tipo;
            SetParameters(cmd, arr);
            int cantidad = cmd.ExecuteNonQuery();
            con.Close();
            return cantidad;
        }

        public object ExecuteScalar(string query, SqlParameter[] arr, CommandType tipo)
        {
            SqlConnection con = GetConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            cmd.CommandType = tipo;
            SetParameters(cmd, arr);
            object valor = cmd.ExecuteScalar();
            con.Close();
            return valor;
        }

        public string ExecuteNonQueryS(string query, SqlParameter[] arr, CommandType tipo)
        {
            SqlConnection con = GetConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            cmd.CommandType = tipo;
            SetParameters(cmd, arr);
            string sRes = cmd.ExecuteNonQuery().ToString();
            con.Close();
            return sRes;
        }

        public DataSet ExecuteDataset(string query, SqlParameter[] arr, CommandType tipo)
        {
            SqlConnection con = GetConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = query;
            cmd.CommandType = tipo;
            SetParameters(cmd, arr);
            SqlDataAdapter adap = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            adap.Fill(ds);
            adap.Dispose();
            con.Close();
            return ds;
        }

        public static string Encriptar(string textoQueEncriptaremos)
        {
            return Encriptar(textoQueEncriptaremos,
              "passc0l0mbi4g0ld", "c0l0mbi4g0ld", "MD5", 1, "@1B2c3D4e5F6g7H8", 128);
        }

        public static string Encriptar(string textoQueEncriptaremos,
        string passBase, string saltValue, string hashAlgorithm,
        int passwordIterations, string initVector, int keySize)
        {
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(textoQueEncriptaremos);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passBase,
              saltValueBytes, hashAlgorithm, passwordIterations);
            byte[] keyBytes = password.GetBytes(keySize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged()
            {
                Mode = CipherMode.CBC
            };
            ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes,
              initVectorBytes);
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor,
             CryptoStreamMode.Write);
            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
            cryptoStream.FlushFinalBlock();
            byte[] cipherTextBytes = memoryStream.ToArray();
            memoryStream.Close();
            cryptoStream.Close();
            string cipherText = Convert.ToBase64String(cipherTextBytes);
            return cipherText;
        }


        /// <summary>
        /// Método para desencriptar un texto encriptado.
        /// </summary>
        /// <returns>Texto desencriptado</returns>
        public static string Desencriptar(string textoEncriptado)
        {
            return Desencriptar(textoEncriptado,
                "passc0l0mbi4g0ld", "c0l0mbi4g0ld", "SHA1", 1, "@1B2c3D4e5F6g7H8", 128);
        }
        /// <summary>
        /// Método para desencriptar un texto encriptado (Rijndael)
        /// </summary>
        /// <returns>Texto desencriptado</returns>
        public static string Desencriptar(string textoEncriptado, string passBase,
          string saltValue, string hashAlgorithm, int passwordIterations,
          string initVector, int keySize)
        {
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);
            byte[] cipherTextBytes = Convert.FromBase64String(textoEncriptado);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passBase,
              saltValueBytes, hashAlgorithm, passwordIterations);
            byte[] keyBytes = password.GetBytes(keySize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged()
            {
                Mode = CipherMode.CBC
            };
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes,
              initVectorBytes);
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
            CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor,
              CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];
            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0,
              plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            string plainText = Encoding.UTF8.GetString(plainTextBytes, 0,
              decryptedByteCount);
            return plainText;
        }

        
    }
}
