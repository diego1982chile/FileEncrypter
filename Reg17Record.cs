using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reg17Generator
{
    class Reg17Record
    {
        String patchCode;
        String publicationDate;
        String product;
        String classification;
        String enhancementsAndCorrections;
        String impactOpinion;

        public Reg17Record(string patchCode, string publicationDate, string product, string classification, string enhancementsAndCorrections, string impactOpinion)
        {
            this.patchCode = patchCode;
            this.PublicationDate = publicationDate;
            this.Product = product;
            this.Classification = classification;
            this.EnhancementsAndCorrections = enhancementsAndCorrections;
            this.ImpactOpinion = impactOpinion;
        }

        public string PatchCode { get => patchCode; set => patchCode = value; }
        public string PublicationDate { get => publicationDate; set => publicationDate = value; }
        public string Product { get => product; set => product = value; }
        public string Classification { get => classification; set => classification = value; }
        public string EnhancementsAndCorrections { get => enhancementsAndCorrections; set => enhancementsAndCorrections = value; }
        public string ImpactOpinion { get => impactOpinion; set => impactOpinion = value; }
    }
}
