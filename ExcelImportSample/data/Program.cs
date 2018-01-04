using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImportSample.data
{
    public class Program
    {
        public Program()
        {
            Remark = "";
        }

        public string ChannelId { get; set; }

        public string Name { get; set; }

        public string AbbreviationName { get; set; }

        public string Kind { get; set; }

        public string DateKind { get; set; }

        public string RelationId { get; set; }

        public DateTime OnAirStart { get; set; }

        public DateTime OnAirEnd { get; set; }

        public string Detail { get; set; }

        public string Remark { get; set; }

        public void SetOnAirStart(string myDate)
        {
            try
            {
                OnAirStart = Convert.ToDateTime(myDate);
            }
            catch (Exception)
            {
                OnAirStart = new DateTime(1900, 1, 1);
                Remark = Remark + "    Error : " + myDate;
            }
        }

        public void SetOnAirEnd(string myDate)
        {
            try
            {
                OnAirEnd = Convert.ToDateTime(myDate);
            }
            catch (Exception)
            {
                OnAirEnd = new DateTime(1900, 1, 1);
                Remark = Remark + "    Error : " + myDate;
            }
        }
    }
}
