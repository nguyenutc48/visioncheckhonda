using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace SaGiangVisionManager
{
    public class SaveXML
    {
        public static void SaveData(object obj, string fileName)
        {
            XmlSerializer sr = new XmlSerializer(obj.GetType());
            TextWriter txtWriter = new StreamWriter(fileName);
            sr.Serialize(txtWriter, obj);
            txtWriter.Close();
        }
    }
}
