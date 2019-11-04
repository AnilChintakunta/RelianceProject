using Newtonsoft.Json;

namespace Reliance
{
    public class PolicyDetailsModel
    {
        [JsonProperty("number")]
        public string Number { get; set; }
        
        [JsonProperty("companyName")]
        public string CompanyName { get; set; }
        
        [JsonProperty("insured")]
        public string Insured { get; set; }
        
        [JsonProperty("address")]
        public string Address { get; set; }
        
        [JsonProperty("pincode")]
        public string Pincode { get; set; }
        
        [JsonProperty("phone")]
        public string Phone { get; set; }
        
        [JsonProperty("email")]
        public string Email { get; set; }
        
        [JsonProperty("registrationNumber")]
        public string RegistrationNumber { get; set; }
        
        [JsonProperty("engineNumber")]
        public string EngineNumber { get; set; }
        
        [JsonProperty("chassis")]
        public string Chassis { get; set; }
        
        [JsonProperty("make")]
        public string Make { get; set; }
        
        [JsonProperty("model")]
        public string Model { get; set; }
        
        [JsonProperty("typeOfBody")]
        public string TypeOfBody { get; set; }
        
        [JsonProperty("yearOfManufactoring")]
        public string YearOfManufacturing { get; set; }
        
        [JsonProperty("seatingCapacity")]
        public string SeatingCapacity { get; set; }
        
        [JsonProperty("vehicleCategory")]
        public string VehicleCategory { get; set; }
        
        [JsonProperty("previousPolicyRSD")]
        public string PreviousPolicyRSD { get; set; }
        
        [JsonProperty("previousPolicyRED")]
        public string PreviousPolicyRED { get; set; }
        
        [JsonProperty("salutation")]
        public string Salutation { get; set; }
        
        [JsonProperty("hypothecation")]
        public string Hypothecation { get; set; }
        
        [JsonProperty("rto")]
        public string RTO { get; set; }
        
        [JsonProperty("ncb")]
        public string NCB { get; set; }
        
        [JsonProperty("nomineeName")]
        public string NomineeName { get; set; }
        
        [JsonProperty("fuelType")]
        public string FuelType { get; set; }
    }
}
