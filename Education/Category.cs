using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;

namespace Education
{
    class Category
    {

        private Dictionary<String,int> mainCategory;
        private ListDictionary subCategories;

        public Category(Dictionary<String,int> mainCategory, ListDictionary subCategories)
        {
            this.mainCategory = mainCategory;
            this.subCategories = subCategories;
        }

        public Dictionary<String,int> MainCategory
        {
            get { return mainCategory; }
            set { mainCategory = value; }
        }
        public ListDictionary SubCategories
        {
            get { return subCategories; }
            set { subCategories = value; }
        }
    }
}
