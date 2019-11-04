using Newtonsoft.Json;

namespace Reliance
{
    public class CustomerDetailsModel
    {
        [JsonProperty(PropertyName = "Policy Number")]
        public string PolicyNumber { get; set; }

        [JsonProperty(PropertyName = "Insured Name")]
        public string InsuredName { get; set; }

        [JsonProperty(PropertyName = "Address")]
        public string Address { get; set; }

        [JsonProperty(PropertyName = "Phone")]
        public string Phone { get; set; }

        [JsonProperty(PropertyName = "Email")]
        public string Email { get; set; }

        [JsonProperty(PropertyName = "Registration Number")]
        public string RegistrationNumber { get; set; }

        [JsonProperty(PropertyName = "Engine Number")]
        public string EngineNumber { get; set; }

        [JsonProperty(PropertyName = "Chassis")]
        public string Chassis { get; set; }

        [JsonProperty(PropertyName = "Make/Model")]
        public string Make_Model { get; set; }

        [JsonProperty(PropertyName = "Type of Body")]
        public string TypeOfBody { get; set; }

        [JsonProperty(PropertyName = "Year of Manufacture")]
        public string YearOfManufacture { get; set; }

        [JsonProperty(PropertyName = "Seating Capacity")]
        public string SeatingCapacity { get; set; }

        [JsonProperty(PropertyName = "Previous policy year")]
        public string PreviousPolicyYear { get; set; }

        [JsonProperty(PropertyName = "Previous policy company name")]
        public string PreviousPolicyCompanyName { get; set; }

        [JsonProperty(PropertyName = "Vehicle Category")]
        public string VehicleCategory { get; set; }

        [JsonProperty(PropertyName = "Vehicle Sub-Category")]
        public string VehicleSubCategory { get; set; }

        [JsonProperty(PropertyName = "Previous Policy RED")]
        public string PreviousPolicyRED { get; set; }

        [JsonProperty(PropertyName = "Salutation")]
        public string Salutation { get; set; }

        [JsonProperty(PropertyName = "Hypothecation")]
        public string Hypothecation { get; set; }

        [JsonProperty(PropertyName = "RTO")]
        public string RTO { get; set; }

        [JsonProperty(PropertyName = "NCB/NCD")]
        public string NCB { get; set; }

    }
}
