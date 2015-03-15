using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Word.Extensions
{
    public static class DocxTemplate
    {
        public static void CriarNovoDocumento(
            string caminhoTemplate,
            string caminhoArquivoDestino,
            Dictionary<string, string> substituicoes)
        {
            File.Copy(caminhoTemplate, caminhoArquivoDestino);

            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(caminhoArquivoDestino, true))
            {
                string docText = null;
                using (StreamReader sr =
                    new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText;
                foreach (var itemSubstituicao in substituicoes)
                {
                    regexText = new Regex(itemSubstituicao.Key);
                    docText = regexText.Replace(docText, itemSubstituicao.Value);
                }

                using (StreamWriter sw = new StreamWriter(
                    wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}