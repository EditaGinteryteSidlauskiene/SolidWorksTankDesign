using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Services;
using System;

namespace SolidWorksTankDesign.MVP.Models
{
    public class NozzleModel : INozzleModel
    {
        private readonly INozzleSolidWorksService _nozzleSWService;

        // Service injection - if present, model delegates SolidWorks work to the service.
        public NozzleModel(INozzleSolidWorksService swService)
        {
            _nozzleSWService = swService ?? throw new ArgumentNullException(nameof(swService));
        }
        
        public void ApplyNozzleChanges()
        {
            if (_nozzleSWService == null) throw new InvalidOperationException("NozzleSolidWorksService not provided.");

            _nozzleSWService.ApplyNozzleChanges();

            //Compartment compartment = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0];
            //compartment.ActivateDocument();
            //Nozzle nozzle = compartment.Nozzles[0];
            ////nozzle.ChangeCutDiameterOfTankBodyEnvelope();
            //nozzle.ChangeTankBodyEnvelopeDimensions(400, 400, 400, 400);
            //DocumentManager.UpdateAndSaveDocuments();
        }
    }
}
