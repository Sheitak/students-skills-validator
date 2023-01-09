using System.Xml.Serialization;

namespace students_skills_validator.Models
{
    public class StudentFile
    {
        public string FirstName { get; set; }

        public string LastName { get; set; }

        [XmlAttribute(AttributeName = "FileName")]
        public string FileName { get; set; }

        public string RefName { get; set; }

        public DateTime updatedAt { get; set; }

        public StudentFile()
        {

        }

        public StudentFile(string FileName)
        {
            this.FileName = FileName;
        }
    }
}
