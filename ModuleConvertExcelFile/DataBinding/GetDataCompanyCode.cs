using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.DataBinding
{
    class GetDataCompanyCode
    {
        private static GetDataCompanyCode instance;

        public static GetDataCompanyCode Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new GetDataCompanyCode();
                }

                return instance;
            }

            set { instance = value; }
        }

        private GetDataCompanyCode() { }
        public DataBindingCompanyCode Get(List<DataBindingCompanyCode> dataBindingCompanyCodes,string CDVHeader, string CDVBody)
        {

            foreach (DataBindingCompanyCode item in dataBindingCompanyCodes)
            {
                if(item.CDVJMemberNumber == (CDVHeader + CDVBody))
                {
                    DataBindingCompanyCode dataBindingCompanyCode = new DataBindingCompanyCode()
                    {
                        CDVJMemberNumber = item.CDVJMemberNumber,
                        GEOStoreCode = item.GEOStoreCode,
                        StoreName = item.StoreName,
                    };
                    return dataBindingCompanyCode;
                }
            }
            return null;
        }
    }
}
