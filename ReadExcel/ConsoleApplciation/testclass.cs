using Xunit;
using System.Collections.Generic;

namespace ConsoleApplication
{

    public class testclass
    {
        [Fact]
        public void TestSubrecipientsCount()
        {
            // Arrange
            
            string filePath = "/Users/jingfei_zhang/Desktop/awardSubawardBudgetExample1.xlsx";
           // NET7ConsoleApp reader = new NET7ConsoleApp(filePath);

            // Act
            List<string> subrecipients = Program.ReadSubrecipients(filePath);
           

            // Assert
            Assert.Equal(4, subrecipients.Count);
            // Assert that the list contains the string ""
           Assert.Contains("Indiana", subrecipients);
           Assert.Contains("Mayo", subrecipients);
           Assert.Contains("Purdue", subrecipients);
           Assert.Contains("Florida", subrecipients);
          
        }
    }
}
