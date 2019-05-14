// --------------------------------------------------------------------------------------------------------------------
// <copyright company="otherslikeyou.com Inc." file="SelectedPage.cs">
//   Licensed under the MIT License. See LICENSE file in the project root for full license information.
// </copyright>
// <summary>
//   
// </summary>
// 
// --------------------------------------------------------------------------------------------------------------------
namespace OLY.OneNotePicker.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class SelectedPage
    {
        public string Title { get; set; }

        public DateTime Modified { get; set; }

        public string Content { get; set; }

        public string SemanticContent => this.GetSemanticContent();

        private string GetSemanticContent()
        {
            return this.Content;
        }
    }
}
