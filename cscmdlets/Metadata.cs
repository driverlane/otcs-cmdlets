using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using cscmdlets.DocumentManagement;

namespace cscmdlets
{
    internal class MetadataManager
    {

        #region globals and constructors

        internal Metadata Metadata
        {
            get { return _Metadata; }
            set { _Metadata = value; }
        }
        private Metadata _Metadata;

        internal MetadataManager()
        {
        }

        internal MetadataManager(Metadata SourceMetadata)
        {
            _Metadata = SourceMetadata;
        }

        #endregion

        #region visible methods

        /// <summary>
        /// Removes all null values from attribute groups on the class metadata object
        /// </summary>
        internal void Clean()
        {
            _Metadata = CleanMetadata(_Metadata);
        }

        internal void AddAttribute(AttributeGroup CategoryTemplate, String Attribute, Object[] Values)
        {

            // create the attribute
            DataValue newAtt = CreateDataValue(CategoryTemplate, Attribute, Values);

            // create the category
            AttributeGroup newCat = ChangeCategoryValue(CategoryTemplate, false, newAtt);

            // add the category
            AddCategory(newCat, true, true);

        }

        /*
        internal void UpdateAttribute(AttributeGroup CategoryTemplate, String Attribute, Dictionary<String, List<Object>> Constraints)
        {

        }
        */

        internal void ClearAttribute(AttributeGroup CategoryTemplate, String Attribute, Dictionary<String, List<Object>> Constraints)
        {
            if (MetadataHasCategory(_Metadata, CategoryTemplate.Key))
            {
                _Metadata = ChangeMetadataValue(_Metadata, true, null, CategoryTemplate.Key);
            }
        }

        internal void AddCategory(AttributeGroup Category, Boolean MergeAttributes, Boolean MergeAttributesUsesNewValues)
        {
            Boolean CatAdded = false;

            // remove the version from the category key
            String catKey = Category.Key.Substring(0, Category.Key.IndexOf("."));

            // add the category, if it's on the metadata object already
            List<AttributeGroup> newCats = new List<AttributeGroup>();
            for (int i = 0; i < _Metadata.AttributeGroups.Length; i++)
            {
                if (_Metadata.AttributeGroups[i].Key.StartsWith(catKey))
                {
                    // check if we're merging the attribute values
                    if (MergeAttributes)
                    {
                        newCats.Add(MergeCategory(_Metadata.AttributeGroups[i], Category, MergeAttributesUsesNewValues));
                        CatAdded = true;
                    }
                    else
                    {
                        newCats.Add(_Metadata.AttributeGroups[i]);
                        CatAdded = true;
                    }
                }
                else
                {
                    newCats.Add(_Metadata.AttributeGroups[i]);
                }
            }
            
            // add the new category if it hasn't been added above
            if (!CatAdded)
            {
                newCats.Add(Category);
            }

            // replace the categories on the metadata object
            _Metadata.AttributeGroups = newCats.ToArray();

        }

        internal void RemoveCategory(Int64 CategoryID)
        {
            List<AttributeGroup> newCats = new List<AttributeGroup>();
            for (int i = 0; i < _Metadata.AttributeGroups.Length; i++)
            {
                if (!_Metadata.AttributeGroups[i].Key.StartsWith(CategoryID.ToString()))
                {
                    newCats.Add(_Metadata.AttributeGroups[i]);
                }
            }
            _Metadata.AttributeGroups = newCats.ToArray();
        }

        internal void AddCategories(List<AttributeGroup> Categories, Boolean MergeCategories, Boolean MergeAttributes, Boolean MergeAttributesUsesNewValues)
        {
            if (MergeCategories)
            {
                foreach (AttributeGroup cat in Categories)
                {
                    AddCategory(cat, MergeAttributes, MergeAttributesUsesNewValues);
                }
            }
            else
            {
                _Metadata.AttributeGroups = Categories.ToArray();
            }
        }

        internal void RemoveCatgories()
        {
            _Metadata.AttributeGroups = null;
        }

        #endregion

        #region private methods

        private Metadata CleanMetadata(Metadata thisMetadata)
        {
            thisMetadata.AttributeGroups = CleanCategories(thisMetadata.AttributeGroups);
            return thisMetadata;
        }

        private AttributeGroup[] CleanCategories(AttributeGroup[] Categories)
        {
            List<AttributeGroup> newCats = new List<AttributeGroup>();
            for (int i = 0; i < Categories.Length; i++)
            {
                newCats.Add(CleanCategory(Categories[i]));
            }
            return newCats.ToArray();
        }

        private AttributeGroup CleanCategory(AttributeGroup Category)
        {

            // create a list to hold the populated attributes
            List<DataValue> newAtts = new List<DataValue>();

            // iterate the category to find atts with values
            for (int i = 0; i < Category.Values.Length; i++)
            {
                // iterate again if it's a set
                if (Category.Values[i].GetType().Equals(typeof(TableValue)))
                {
                    // cast the attribute to a table value to access the rows array
                    TableValue setAtt = (TableValue)Category.Values[i];

                    // create a list of rows to store the populated attributes
                    List<RowValue> rows = new List<RowValue>();

                    // iterate the rows to find atts with values
                    for (int j = 0; j < setAtt.Values.Length; j++)
                    {
                        // iterate the atts on the row
                        List<DataValue> rowAtts = new List<DataValue>();
                        for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                        {
                            // check the set row level attribute
                            if (AttributeHasValues(setAtt.Values[j].Values[k]))
                            {
                                rowAtts.Add(setAtt.Values[j].Values[k]);
                            }
                        }

                        // create the row object if there's any values
                        if (rowAtts.Count > 0)
                        {
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
                    if (AttributeHasValues(Category.Values[i]))
                    {
                        newAtts.Add(Category.Values[i]);
                    }
                }

                // update the atts on the category
                Category.Values = newAtts.ToArray();

            }

            return Category;

        }

        private Boolean AttributeHasValues(DataValue Attribute)
        {

            // check if there are values on the real attribute type
            if (Attribute.GetType().Equals(typeof(BooleanValue)))
            {
                BooleanValue boolAtt = (BooleanValue)Attribute;
                if (boolAtt.Values.Length > 0) return true;
            }
            else if (Attribute.GetType().Equals(typeof(DateValue)))
            {
                DateValue dateAtt = (DateValue)Attribute;
                if (dateAtt.Values.Length > 0) return true;
            }
            else if (Attribute.GetType().Equals(typeof(IntegerValue)))
            {
                IntegerValue intAtt = (IntegerValue)Attribute;
                if (intAtt.Values.Length > 0) return true;
            }
            else if (Attribute.GetType().Equals(typeof(RealValue)))
            {
                RealValue realAtt = (RealValue)Attribute;
                if (realAtt.Values.Length > 0) return true;
            }
            else if (Attribute.GetType().Equals(typeof(StringValue)))
            {
                StringValue strAtt = (StringValue)Attribute;
                if (strAtt.Values.Length > 0) return true;
            }

            throw new Exception("Unsupported attribute data type.");

        }

        private DataValue CreateDataValue(AttributeGroup CategoryTemplate, String Attribute, Object[] Values)
        {
            return (DataValue)GetAttribute(CategoryTemplate, Attribute, false, Values);
        }

        private Object GetAttribute(AttributeGroup CategoryTemplate, String Attribute, Boolean ReturnExisting, Object[] Values)
        {

            // find the attribute
            for (int i = 0; i < CategoryTemplate.Values.Length; i++)
            {
                // check a first level attribute
                if (CategoryTemplate.Values[i].Description == Attribute)
                {
                    if (Values != null)
                    {
                        return PopulateDataValue(CategoryTemplate.Values[i], Values);
                    }
                    else if (ReturnExisting)
                    {
                        return CategoryTemplate.Values[i];
                    }
                }

                // recurse any set rows
                if (CategoryTemplate.Values[i].GetType().Equals(typeof(TableValue)))
                {
                    TableValue setAtt = (TableValue)CategoryTemplate.Values[i];
                    for (int j = 0; j < setAtt.Values.Length; j++)
                    {
                        // iterate the atts on the row
                        for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                        {
                            // check the set row level attribute
                            if (setAtt.Values[j].Values[k].Description == Attribute)
                            {
                                if (Values != null)
                                {
                                    return PopulateDataValue(setAtt.Values[j].Values[k], Values);
                                }
                                else if (ReturnExisting)
                                {
                                    return setAtt.Values[j].Values[k];
                                }
                            }
                        }
                    }
                }

            }

            throw new Exception("Attribute not found on specified category");

        }

        private DataValue PopulateDataValue(DataValue Attribute, Object[] Values)
        {

            // return the real attribute type
            if (Attribute.GetType().Equals(typeof(BooleanValue)))
            {
                BooleanValue boolAtt = new BooleanValue();
                boolAtt.Description = Attribute.Description;
                boolAtt.Key = Attribute.Key;
                boolAtt.Values = Values.Cast<Boolean?>().ToArray();
                return boolAtt;
            }
            else if (Attribute.GetType().Equals(typeof(DateValue)))
            {
                DateValue dateAtt = new DateValue();
                dateAtt.Description = Attribute.Description;
                dateAtt.Key = Attribute.Key;
                dateAtt.Values = Values.Cast<DateTime?>().ToArray();
                return dateAtt;
            }
            else if (Attribute.GetType().Equals(typeof(IntegerValue)))
            {
                IntegerValue intAtt = new IntegerValue();
                intAtt.Description = Attribute.Description;
                intAtt.Key = Attribute.Key;
                intAtt.Values = Values.Cast<Int64?>().ToArray();
                return intAtt;
            }
            else if (Attribute.GetType().Equals(typeof(RealValue)))
            {
                RealValue realAtt = new RealValue();
                realAtt.Description = Attribute.Description;
                realAtt.Key = Attribute.Key;
                realAtt.Values = Values.Cast<Double?>().ToArray();
                return realAtt;
            }
            else if (Attribute.GetType().Equals(typeof(StringValue)))
            {
                StringValue strAtt = new StringValue();
                strAtt.Description = Attribute.Description;
                strAtt.Key = Attribute.Key;
                strAtt.Values = Values.Cast<String>().ToArray();
                return strAtt;
            }

            throw new Exception("Unsupported attribute data type.");

        }

        private Metadata ChangeMetadataValue(Metadata Metadata, Boolean ClearValue, DataValue Attribute, String CategoryKey)
        {
            Metadata.AttributeGroups = ChangeCategoriesValue(Metadata.AttributeGroups, ClearValue, Attribute, CategoryKey);
            return Metadata;
        }

        private AttributeGroup[] ChangeCategoriesValue(AttributeGroup[] Categories, Boolean ClearValue, DataValue Attribute, String CategoryKey)
        {
            // remove the version from the category key
            CategoryKey = CategoryKey.Substring(0, CategoryKey.IndexOf("."));

            // create the list for the new categories
            List<AttributeGroup> newCats = new List<AttributeGroup>();

            // find the category
            for (int i = 0; i < Categories.Length; i++)
            {
                if (Categories[i].Key.StartsWith(CategoryKey))
                {
                    newCats.Add(ChangeCategoryValue(Categories[i], ClearValue, Attribute));
                }
                else
                {
                    newCats.Add(Categories[i]);
                }
            }

            // convert the list to array and return
            return newCats.ToArray();
        }

        private AttributeGroup ChangeCategoryValue(AttributeGroup Category, Boolean ClearValue, DataValue Attribute)
        {
            // create the list for the new attributes
            List<DataValue> newAtts = new List<DataValue>();

            // find the attribute
            for (int i = 0; i < Category.Values.Length; i++)
            {

                // recurse any set rows
                if (Category.Values[i].GetType().Equals(typeof(TableValue)))
                {
                    // cast the attribute to a table to get access to the rows array
                    TableValue setAtt = (TableValue)Category.Values[i];

                    // create list for the new rows
                    List<RowValue> newRows = new List<RowValue>();

                    // iterate the rows
                    for (int j = 0; j < setAtt.Values.Length; j++)
                    {
                        // create list for the new row attributes
                        List<DataValue> newRowAtts = new List<DataValue>();

                        // iterate the atts on the row
                        for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                        {
                            // check the set row level attribute
                            if (setAtt.Values[j].Values[k].Description == Attribute.Description)
                            {
                                newRowAtts.Add(Attribute);
                            }
                            else
                            {
                                newRowAtts.Add(setAtt.Values[j].Values[k]);
                            }
                        }

                        // add the row attributes to the row object and that to the rows list
                        RowValue newRow = new RowValue();
                        newRow.Values = newRowAtts.ToArray();
                        newRows.Add(newRow);

                    }

                    // add the rows to the set and add that to the new attributes list
                    setAtt.Values = newRows.ToArray();
                    newAtts.Add(setAtt);
                }

                // check the standard attribute
                else
                {
                    if (Category.Values[i].Description == Attribute.Description)
                    {
                        newAtts.Add(Attribute);
                    }
                    else
                    {
                        newAtts.Add(Category.Values[i]);
                    }
                }

            }

            // add the new attributes to the category and return it
            Category.Values = newAtts.ToArray();
            return Category;
        }

        private AttributeGroup MergeCategory(AttributeGroup OldCategory, AttributeGroup NewCategory, Boolean MergeAttributesUsesNewValues)
        {

            // assume we're using the old values as source
            AttributeGroup catOne = OldCategory;
            AttributeGroup catTwo = NewCategory;

            // change if we're using the new values
            if (MergeAttributesUsesNewValues)
            {
                catOne = NewCategory;
                catTwo = OldCategory;
            }
           
            // the list for the new attributes
            List<DataValue> newAtts = new List<DataValue>();

            // iterate cat one, replacing nulls with cat two values
            for (int i = 0; i < catOne.Values.Length; i++)
            {

                // recurse any set rows
                if (catOne.Values[i].GetType().Equals(typeof(TableValue)))
                {
                    // cast the attribute to a table to get access to the rows array
                    TableValue setAtt = (TableValue)catOne.Values[i];

                    // create list for the new rows
                    List<RowValue> newRows = new List<RowValue>();

                    // iterate the rows
                    for (int j = 0; j < setAtt.Values.Length; j++)
                    {
                        // create list for the new row attributes
                        List<DataValue> newRowAtts = new List<DataValue>();

                        // iterate the atts on the row
                        for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                        {
                            // if it's populated, use it
                            if (AttributeHasValues(setAtt.Values[j].Values[k]))
                            {
                                newRowAtts.Add(setAtt.Values[j].Values[k]);
                            }
                            // it's not populated, so see if the other cat has a value and use that
                            else
                            {
                                DataValue otherAtt = (DataValue)GetAttribute(catTwo, setAtt.Values[j].Values[k].Description, true, null);
                                if (AttributeHasValues(otherAtt))
                                {
                                    newRowAtts.Add(otherAtt);
                                }
                                else
                                {
                                    newRowAtts.Add(setAtt.Values[j].Values[k]);
                                }
                            }
                        }

                        // add the row attributes to the row object and that to the rows list
                        RowValue newRow = new RowValue();
                        newRow.Values = newRowAtts.ToArray();
                        newRows.Add(newRow);

                    }

                    // add the rows to the set and add that to the new attributes list
                    setAtt.Values = newRows.ToArray();
                    newAtts.Add(setAtt);
                }


                // check the standard attribute
                else
                {
                    // if it's populated, use it
                    if (AttributeHasValues(catOne.Values[i]))
                    {
                        newAtts.Add(catOne.Values[i]);
                    }
                    // it's not populated, so see if the other cat has a value and use that
                    else
                    {
                        DataValue otherAtt = (DataValue)GetAttribute(catTwo, catOne.Values[i].Description, true, null);
                        if (AttributeHasValues(otherAtt))
                        {
                            newAtts.Add(otherAtt);
                        }
                        else
                        {
                            newAtts.Add(catOne.Values[i]);
                        }
                    }
                }

            }
            
            // return the old category with the new values
            OldCategory.Values = newAtts.ToArray();
            return OldCategory;

        }

        private Boolean MetadataHasCategory(Metadata Metadata, String CategoryKey)
        {
            // remove the version from the category key
            CategoryKey = CategoryKey.Substring(0, CategoryKey.IndexOf("."));

            // check if the category exists
            for (int i = 0; i < Metadata.AttributeGroups.Length; i++)
            {
                if (Metadata.AttributeGroups[i].Key.StartsWith(CategoryKey)) return true;
            }
            return false;
        }

        #endregion

        /*

        private String GetAttributeID(Int64 CategoryID, String Attribute)
        {
            return (String)GetAttribute(CategoryID, Attribute, null);
        }

        */

    }
}
