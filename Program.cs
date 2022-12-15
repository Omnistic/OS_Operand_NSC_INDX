using System;
using ZOSAPI;
using ZOSAPI.Tools;

namespace CSharpUserOperandApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginUserOperand();
        }

        static void BeginUserOperand()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Operand)
            {
                HandleError("User plugin was started in the wrong mode: expected Operand, found " + TheApplication.Mode.ToString());
                return;
            }

            // Read the operand arguments
            double Hx = TheApplication.OperandArgument1;
            double Hy = TheApplication.OperandArgument2;
            double Px = TheApplication.OperandArgument3;
            double Py = TheApplication.OperandArgument4;

            // Initialize the output array
            int maxResultLength = TheApplication.OperandResults.Length;
            double[] operandResults = new double[maxResultLength];

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;
            // Add your custom code here...

            // Object number from which to evaluate material index
            int object_number = (int)Hx;

            // Wave number at which to evaluate material index
            int wave_number = (int)Hy;

            // Actual wavelength
            double wavelength = TheSystem.SystemData.Wavelengths.GetWavelength(wave_number).Wavelength;

            // Wavelength squared
            double wavel_squared = wavelength * wavelength;

            // Catalog number (following look-up "switch" below, e.g. 1 = Schott)
            int catalog_number = (int)Px;

            // Catalog look-up
            string catalog = "";
            switch(catalog_number)
            {
                case 1:
                    catalog = "Schott";
                    break;
                // Add other catalogs here
                // ...
            }

            // Get Material name
            string material = TheSystem.NCE.GetObjectAt(object_number).Material;

            // Open >> Librairies..Materials Catalog
            IMaterialsCatalog material_cat = TheSystem.Tools.OpenMaterialsCatalog();

            // Select catalog and material
            material_cat.SelectedCatalog = catalog;
            material_cat.SelectedMaterial = material;

            // Run Materials Catalog
            material_cat.RunAndWaitForCompletion();

            // Initialize dummy refractive index value
            double refractive_index = -1.0;

            // Look-up index formula
            switch (material_cat.MaterialFormula)
            {
                // Sellmeier 1
                case MaterialFormulas.Sellmeier1:
                    // Initialize fit coefficients
                    double K1, K2, K3, L1, L2, L3;

                    // Get fit coefficients
                    K1 = material_cat.GetFitCoefficient(0);
                    K2 = material_cat.GetFitCoefficient(2);
                    K3 = material_cat.GetFitCoefficient(4);
                    L1 = material_cat.GetFitCoefficient(1);
                    L2 = material_cat.GetFitCoefficient(3);
                    L3 = material_cat.GetFitCoefficient(5);

                    // Refractive index calculation
                    refractive_index = K1 / (wavel_squared - L1);
                    refractive_index += K2 / (wavel_squared - L2);
                    refractive_index += K3 / (wavel_squared - L3);
                    refractive_index *= wavel_squared;
                    refractive_index += 1.0;
                    refractive_index = Math.Sqrt(refractive_index);

                    break;
                // Add other formulas here
                // ...
            }

            // Close Materials Catalog
            material_cat.Close();

            operandResults[0] = refractive_index;

            // Clean up
            FinishUserOperand(TheApplication, operandResults);
        }

        static void FinishUserOperand(IZOSAPI_Application TheApplication, double[] resultData)
        {
            // Note - OpticStudio will wait for the operand to complete until this application exits 
            if (TheApplication != null)
            {
                TheApplication.OperandResults.WriteData(resultData.Length, resultData);
            }
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
