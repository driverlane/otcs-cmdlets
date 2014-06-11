using cscmdlets.DocumentManagement;
using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;

namespace cscmdlets
{
    public abstract class CSMetadata
    {

        public Metadata AddCategory(Metadata Metadata, AttributeGroup Category, Boolean Replace, Boolean MergeAttributes, Boolean UseNewValues)
        {
            // get a non-version specific category key
            String catKey = Category.Key.Substring(0, Category.Key.IndexOf('.'));

            // build the list of categories and check if we've got the category
            List<AttributeGroup> cats = new List<AttributeGroup>();
            if (Metadata.AttributeGroups != null)
            {
                for (int i = 0; i < Metadata.AttributeGroups.Length; i++)
                {
                    if (Metadata.AttributeGroups[i].Key.StartsWith(catKey))
                    {
                        if (Replace)
                            cats.Add(Category);
                        else if (MergeAttributes)
                            cats.Add(MergeAttributeGroup(Metadata.AttributeGroups[i], Category, UseNewValues));
                        else
                            throw new Exception("Category already on object.");
                    }
                    else
                        cats.Add(Metadata.AttributeGroups[i]);
                }
            }

            // add the new category and reset the metadata object
            cats.Add(Category);
            Metadata.AttributeGroups = cats.ToArray();
            return Metadata;            
        }

        public List<Object> GetAttributeValues(DataValue Attribute)
        {
            List<Object> response = null;

            // check if there are values on the real attribute type
            if (Attribute.GetType().Equals(typeof(BooleanValue)))
            {
                BooleanValue boolAtt = (BooleanValue)Attribute;
                if (boolAtt.Values != null && boolAtt.Values.Length > 0)
                {
                    response = new List<Object>();
                    foreach (bool? val in boolAtt.Values)
                    {
                        response.Add(val);
                    }
                }
                return response;
            }
            else if (Attribute.GetType().Equals(typeof(DateValue)))
            {
                DateValue dateAtt = (DateValue)Attribute;
                if (dateAtt.Values != null && dateAtt.Values.Length > 0)
                {
                    response = new List<Object>();
                    foreach (DateTime? val in dateAtt.Values)
                    {
                        response.Add(val);
                    }
                }
                return response;
            }
            else if (Attribute.GetType().Equals(typeof(IntegerValue)))
            {
                IntegerValue intAtt = (IntegerValue)Attribute;
                if (intAtt.Values != null && intAtt.Values.Length > 0)
                {
                    response = new List<Object>();
                    foreach (long? val in intAtt.Values)
                    {
                        response.Add(val);
                    }
                }
                return response;
            }
            else if (Attribute.GetType().Equals(typeof(RealValue)))
            {
                RealValue realAtt = (RealValue)Attribute;
                if (realAtt.Values != null && realAtt.Values.Length > 0)
                {
                    response = new List<Object>();
                    foreach (double? val in realAtt.Values)
                    {
                        response.Add(val);
                    }
                }
                return response;
            }
            else if (Attribute.GetType().Equals(typeof(StringValue)))
            {
                StringValue strAtt = (StringValue)Attribute;
                if (strAtt.Values != null && strAtt.Values.Length > 0)
                {
                    response = new List<Object>();
                    foreach (string val in strAtt.Values)
                    {
                        response.Add(val);
                    }
                }
                return response;
            }

            throw new Exception("Unsupported attribute data type.");
        }

        public AttributeGroup GetCategory(Metadata Metadata, Int64 CategoryID)
        {
            if (Metadata.AttributeGroups != null)
            {
                for (int i = 0; i < Metadata.AttributeGroups.Length; i++)
                {
                    if (Metadata.AttributeGroups[i].Values != null && Metadata.AttributeGroups[i].Key.StartsWith(CategoryID.ToString()))
                        return Metadata.AttributeGroups[i];
                }
            }
            return null;
        }

        public Metadata CleanMetadata(Metadata Metadata)
        {
            // call the method to clean any categories
            if (Metadata.AttributeGroups != null) Metadata.AttributeGroups = CleanAttributeGroups(Metadata.AttributeGroups);
            return Metadata;
        }

        public AttributeGroup[] CleanAttributeGroups(AttributeGroup[] Categories)
        {
            // call the method to clean any category
            List<AttributeGroup> newCats = new List<AttributeGroup>();
            for (int i = 0; i < Categories.Length; i++)
            {
                if (Categories[i].Values != null) newCats.Add(CleanAttributeGroup(Categories[i]));
            }
            return newCats.ToArray();
        }

        public AttributeGroup CleanAttributeGroup(AttributeGroup Category)
        {
            // create a list to hold the populated attributes
            List<DataValue> newAtts = new List<DataValue>();

            // iterate the category to find atts with values
            if (Category.Values != null)
                for (int i = 0; i < Category.Values.Length; i++)
                {
                    if (Category.Values[i].GetType().Equals(typeof(TableValue)))
                    {
                        // cast the attribute to a table value to access the rows array
                        TableValue setAtt = (TableValue)Category.Values[i];

                        // create a list of rows to store the populated attributes
                        List<RowValue> rows = new List<RowValue>();

                        // iterate the rows to find atts with values
                        if (setAtt.Values != null)
                        {
                            for (int j = 0; j < setAtt.Values.Length; j++)
                            {
                                // iterate the atts on the row
                                List<DataValue> rowAtts = new List<DataValue>();
                                for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                                {
                                    // check the set row level attribute
                                    if (!AttributeIsNullString(setAtt.Values[j].Values[k]))
                                    {
                                        rowAtts.Add(setAtt.Values[j].Values[k]);
                                    }
                                }

                                // create the row object if there's any values
                                RowValue row = new RowValue();
                                row.Description = setAtt.Values[j].Description;
                                row.Key = setAtt.Values[j].Key;
                                row.Values = rowAtts.ToArray();

                                // add the row to the list
                                rows.Add(row);

                            }
                        }

                        // update the set values and add it to the new list
                        setAtt.Values = rows.ToArray();
                        newAtts.Add(setAtt);

                    }

                    // it's not a set, so just check the attribute
                    else
                    {
                        if (!AttributeIsNullString(Category.Values[i]))
                        {
                            newAtts.Add(Category.Values[i]);
                        }
                    }

                }

            // update the atts on the category
            Category.Values = newAtts.ToArray();
            return Category;

        }

        private Boolean AttributeIsNullString(DataValue Attribute)
        {
            // if it's a string value, check we've got some values
            if (Attribute.GetType().Equals(typeof(StringValue)))
            {
                // if there's values return false
                StringValue strAtt = (StringValue)Attribute;
                if (strAtt.Values != null && strAtt.Values.Length > 0)
                    return false;
                else
                    return true;
            }
            // not a string value so return false
            return false;
        }

        private AttributeGroup MergeAttributeGroup(AttributeGroup OldCategory, AttributeGroup NewCategory, Boolean UseNewValues)
        {
            // todo merge the categories
            return null;
        }
    }
}
