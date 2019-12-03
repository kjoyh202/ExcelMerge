using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge
{
    public static class Helper
    {
        public static T SerializableDeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var bformatter = new BinaryFormatter();
                bformatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)bformatter.Deserialize(ms);
            }
        }
    }
}
