using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXml.Word.Extensions;

namespace OpenXml.Word.Tests
{
    [TestClass]
    public class DocxTemplateTeste01
    {
        [TestMethod]
        public void TestarGeracaoDocx()
        {
            Dictionary<string, string> substituicoes =
                new Dictionary<string, string>();
            substituicoes["#NOME_CLIENTE#"] =
                "João da Silva";
            substituicoes["#ENDERECO_CLIENTE#"] =
                "Avenida Paulista, 950 - São Paulo - SP";
            substituicoes["#NOME_ASSINATURA#"] =
                "Pedro Oliveira";

            string caminhoTemplate =
                ConfigurationManager.AppSettings["CaminhoArquivoTemplate"];
            string caminhoArquivoDestino =
                ConfigurationManager.AppSettings["DiretorioGeracaoArquivoTeste"] +
                "Teste_" + DateTime.Now.ToString("dd-MM-yyyy_HH'h'mm'min'ss's'") + ".docx";

            DocxTemplate.CriarNovoDocumento(
                caminhoTemplate,
                caminhoArquivoDestino,
                substituicoes);

            Assert.IsTrue(File.Exists(caminhoArquivoDestino));
        }
    }
}
