/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
    sealed class NumberingListStyleCollection
    {
        public const string HEADING_NUMBERING_NAME = "decimal-heading-multi";
        const string OrderingTypeDecimal = "decimal";
        const string OrderingTypeDisc = "disc";
        const string OrderingTypeSquare = "square";
        const string OrderingTypeCircle = "circle";
        const string OrderingTypeUpperAlpha = "upper-alpha";
        const string OrderingTypeLowerAlpha = "lower-alpha";
        const string OrderingTypeUpperRoman = "upper-roman";
        const string OrderingTypeLowerRoman = "lower-roman";

        private MainDocumentPart mainPart;
        private int nextInstanceID;
        private int levelDepth;
        private int maxlevelDepth;
        private bool firstItem;
        private readonly Stack<KeyValuePair<int, int>> numInstances = new Stack<KeyValuePair<int, int>>();
        private readonly Stack<string[]> listHtmlElementClasses = new Stack<string[]>();
        private int headingNumberingId;

        #region Constructor

        public NumberingListStyleCollection(MainDocumentPart mainPart)
        {
            this.mainPart = mainPart;
            InitNumberingIds();
        }

        #endregion

        #region InitNumberingIds

        private AbstractNum[] CreateDefaultNumberings(int absNumIdRef)
        {
            var defaultItems = new[] {
				//8 kinds of abstractnum + 1 multi-level.
				new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeDecimal } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "•" },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 1, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeDisc } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "▪" },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 2, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeSquare } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "o" },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 3, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeCircle } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 4, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeUpperAlpha } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 5, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeLowerAlpha } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." },
                        //LevelRestart = new LevelRestart(),
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 6, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeUpperRoman } },
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." },
                        // LevelRestart = new LevelRestart(),                           
                        PreviousParagraphProperties = new PreviousParagraphProperties {
                            Indentation = new Indentation() { Left = "420", Hanging = "360" }
                        }
                    }
                ) { AbstractNumberId = absNumIdRef + 7, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = OrderingTypeLowerRoman } },
				// decimal-heading-multi
				// WARNING: only use this for headings
				new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                    new Level {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        LevelIndex = 0,
                        LevelText = new LevelText() { Val = "%1." }
                    }
                ) { AbstractNumberId = absNumIdRef + 8, AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = HEADING_NUMBERING_NAME } }
            };

            return defaultItems;
        }

        private void InitNumberingIds()
        {
            NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart;
            int absNumIdRef = 0;

            // Ensure the numbering.xml file exists or any numbering or bullets list will results
            // in simple numbering list (1.   2.   3...)
            if (numberingPart == null)
                numberingPart = numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();

            if (mainPart.NumberingDefinitionsPart.Numbering == null)
            {
                new Numbering().Save(numberingPart);
            }
            else
            {
                // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
                // to retrieve the highest Id and reconstruct our own list definition template.
                foreach (var abs in numberingPart.Numbering.Elements<AbstractNum>())
                {
                    if (abs.AbstractNumberId.HasValue && abs.AbstractNumberId > absNumIdRef)
                        absNumIdRef = abs.AbstractNumberId;
                }
                absNumIdRef++;
            }

            // This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
            // http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en

            AbstractNum[] absNumChildren = CreateDefaultNumberings(absNumIdRef);

            // Check if we have already initialized our abstract nums
            // if that is the case, we should not add them again.
            // This supports a use-case where the HtmlConverter is called multiple times
            // on document generation, and needs to continue existing lists
            var addNewAbstractNums = false;
            var existingAbstractNums = AbstractNums;

            if (existingAbstractNums.Count() >= absNumChildren.Length) // means we might have added our own already
            {
                foreach (var abstractNum in absNumChildren)
                {
                    // Check if we can find this in the existing document
                    addNewAbstractNums = addNewAbstractNums
                       || !existingAbstractNums.Any(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val.Value == abstractNum.AbstractNumDefinitionName.Val.Value);
                }
            }
            else
            {
                addNewAbstractNums = true;
            }

            if (addNewAbstractNums)
            {
                // this is not documented but MS Word needs that all the AbstractNum are stored consecutively.
                // Otherwise, it will apply the "NoList" style to the existing ListInstances.
                // This is the reason why I insert all the items after the last AbstractNum.
                int lastAbsNumIndex = 0;
                if (absNumIdRef > 0)
                {
                    lastAbsNumIndex = numberingPart.Numbering.ChildElements.Count - 1;
                    for (; lastAbsNumIndex >= 0; lastAbsNumIndex--)
                    {
                        if (numberingPart.Numbering.ChildElements[lastAbsNumIndex] is AbstractNum)
                            break;
                    }
                }

                lastAbsNumIndex = lastAbsNumIndex == -1 ? 0 : lastAbsNumIndex;

                for (int i = 0; i < absNumChildren.Length; i++)
                    numberingPart.Numbering.InsertAt(absNumChildren[i], i + lastAbsNumIndex);
            }

            // compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
            // The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
            // at this level of the style hierarchy. While processing this markup, if the w:val='0',
            // the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
            nextInstanceID = 1;
            foreach (NumberingInstance inst in numberingPart.Numbering.Elements<NumberingInstance>())
            {
                if (inst.NumberID.Value > nextInstanceID) nextInstanceID = inst.NumberID;
            }
            numInstances.Push(new KeyValuePair<int, int>(nextInstanceID, -1));

            numberingPart.Numbering.Save();
        }

        #endregion

        #region BeginList

        public void BeginList(HtmlEnumerator en)
        {
            // lookup for a predefined list style in the template collection
            var type = en.StyleAttributes["list-style-type"];
            var orderedList =
                en.CurrentTag.Equals("<ol>", StringComparison.OrdinalIgnoreCase)
                || OrderedTypes.Contains(type.ToLowerInvariant());

            CreateList(type, orderedList);
            listHtmlElementClasses.Push(en.Attributes.GetAsClass());
        }

        #endregion

        #region EndList

        public void EndList(bool forcePopInstances = true)
        {
            levelDepth--;
            firstItem = true;//levelDepth == 0;

            //var popInstances = levelDepth > 0 || forcePopInstances;
            var popInstances = forcePopInstances;
            if (popInstances)
            {
                numInstances.Pop();  // decrement for nested list
            }


            if (listHtmlElementClasses.Any())
            {
                listHtmlElementClasses.Pop();
            }


            Console.WriteLine($"EndList levelDepth {levelDepth}\tpopInstances {popInstances}\tforcePopInstances {forcePopInstances}");
        }

        #endregion

        #region SetLevelDepth

        public void SetLevelDepth(int newLevelDepth)
        {
            levelDepth = newLevelDepth;
        }

        #endregion

        #region Headings

        public int GetHeadingNumberingId()
        {
            if (headingNumberingId == default(int))
            {
                int absNumberId = GetAbstractNumberIdFromType(HEADING_NUMBERING_NAME, true).AbstractNumberId.Value;

                var existingTitleNumbering = mainPart.NumberingDefinitionsPart.Numbering
                    .Elements<NumberingInstance>()
                    .FirstOrDefault(n => n != null && n.AbstractNumId.Val == absNumberId);

                if (existingTitleNumbering != null)
                {
                    headingNumberingId = existingTitleNumbering.NumberID.Value;
                }
                else
                {
                    headingNumberingId = CreateList(HEADING_NUMBERING_NAME, true);
                    EnsureMultilevel(absNumberId, true);
                }
            }

            Console.WriteLine($"GetHeadingNumberingId() returns {headingNumberingId}");

            return headingNumberingId;
        }

        public void ApplyNumberingToHeadingParagraph(Paragraph p, int indentLevel)
        {
            Console.WriteLine($"ApplyNumberingToHeadingParagraph indentLevel {indentLevel - 1}");

            // Apply numbering to paragraph
            p.InsertInProperties(prop => prop.NumberingProperties = new NumberingProperties(
                new NumberingLevelReference() { Val = (indentLevel - 1) }, // indenting starts at 0
                new NumberingId() { Val = GetHeadingNumberingId() }
            ));

            // Make sure we reset everything for upcoming lists
            EndList(false);
            SetLevelDepth(0);
        }

        #endregion

        #region CreateList

        public int CreateList(string type, bool orderedList)
        {
            var abstractNumber = GetAbstractNumberIdFromType(type, orderedList);
            int absNumId = abstractNumber.AbstractNumberId.Value;
            int prevAbsNumId = numInstances.Peek().Value;

            firstItem = true;
            levelDepth++;
            if (levelDepth > maxlevelDepth)
            {
                maxlevelDepth = levelDepth;
            }

            // save a NumberingInstance if the nested list style is the same as its ancestor.
            // this allows us to nest <ol> and restart the indentation to 1.
            int currentInstanceId = InstanceId;
            if (levelDepth > 1 && absNumId == prevAbsNumId && orderedList)
            {
                EnsureMultilevel(absNumId);
            }
            else
            {
                // For unordered lists (<ul>), create only one NumberingInstance per level
                // (MS Word does not tolerate hundreds of identical NumberingInstances)
                if (orderedList || (levelDepth >= maxlevelDepth))
                {
                    currentInstanceId = ++nextInstanceID;
                    var numbering = mainPart.NumberingDefinitionsPart.Numbering;

                    numbering.Append(new NumberingInstance(
                            new AbstractNumId() { Val = absNumId },
                            new LevelOverride(new StartOverrideNumberingValue() { Val = 1 }) { LevelIndex = 0, }
                        )
                    { NumberID = currentInstanceId, });

                    numbering.Save(mainPart.NumberingDefinitionsPart);
                    mainPart.NumberingDefinitionsPart.Numbering.Reload();
                }
            }

            numInstances.Push(new KeyValuePair<int, int>(currentInstanceId, absNumId));

            Console.WriteLine($"BeginList levelDepth {levelDepth} / {currentInstanceId} - {absNumId}");

            return currentInstanceId;
        }

        #endregion

        #region GetAbstractNumberIdFromType

        public AbstractNum GetAbstractNumberIdFromType(string type)
        {
            return AbstractNums
                     .Where(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val != null)
                     .FirstOrDefault(x => x.AbstractNumDefinitionName.Val.Value == type?.ToLowerInvariant())
                     ;
            //.ToDictionary(a => a.AbstractNumDefinitionName.Val.Value, a => a.AbstractNumberId.Value);
        }

        public AbstractNum GetAbstractNumberIdFromType(string type, bool orderedList)
        {
            var knownAbsNumIds = GetAbstractNumberIdFromType(type);

            if (type == null || knownAbsNumIds == null)
            {
                if (orderedList)
                    knownAbsNumIds = GetAbstractNumberIdFromType("decimal");
                else
                    knownAbsNumIds = GetAbstractNumberIdFromType("disc");
            }

            return knownAbsNumIds;
        }

        #endregion

        #region ProcessItem

        public int ProcessItem(HtmlEnumerator en)
        {
            //Console.WriteLine($"ProcessItem en {en.}");

            if (!firstItem)
            {
                return InstanceId;
            }

            firstItem = false;

            // in case a margin has been specifically specified, we need to create a new list template
            // on the fly with a different AbsNumId, in order to let Word doesn't merge the style with its predecessor.
            Margin margin = en.StyleAttributes.GetAsMargin("margin");
            if (margin.Left.Value > 0 && margin.Left.Type == UnitMetric.Pixel)
            {
                CreateNewLevel();
            }

            return InstanceId;
        }

        private void CreateNewLevel()
        {
            var absNumId = numInstances.Peek().Value;
            var absNum = AbstractNums.FirstOrDefault(a => a.AbstractNumberId.Value == absNumId);

            if (absNum != null)
            {
                var lvl = absNum.GetFirstChild<Level>();
                var currentNumId = ++nextInstanceID;

                var numbering = mainPart.NumberingDefinitionsPart.Numbering;
                numbering.Append(
                    new AbstractNum(
                            new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                            new Level
                            {
                                StartNumberingValue = new StartNumberingValue() { Val = 1 },
                                NumberingFormat = new NumberingFormat() { Val = lvl.NumberingFormat.Val },
                                LevelIndex = 0,
                                LevelRestart = new LevelRestart(),
                                LevelText = new LevelText() { Val = lvl.LevelText.Val }
                            }
                        )
                    { AbstractNumberId = currentNumId });
                numbering.Save(mainPart.NumberingDefinitionsPart);

                numbering.Append(new NumberingInstance(new AbstractNumId() { Val = currentNumId }) { NumberID = currentNumId });
                
                numbering.Save(mainPart.NumberingDefinitionsPart);
                mainPart.NumberingDefinitionsPart.Numbering.Reload();
            }

        }

        #endregion

        #region EnsureMultilevel

        /// <summary> Find a specified AbstractNum by its ID and update its definition to make it multi-level. </summary>
        private void EnsureMultilevel(int absNumId, bool cascading = false)
        {
            var absNumMultilevel = AbstractNums.SingleOrDefault(a => a.AbstractNumberId.Value == absNumId);

            if (absNumMultilevel != null && absNumMultilevel.MultiLevelType.Val == MultiLevelValues.SingleLevel)
            {
                var level1 = absNumMultilevel.GetFirstChild<Level>();
                absNumMultilevel.MultiLevelType.Val = MultiLevelValues.Multilevel;

                // skip the first level, starts to 2
                for (var i = 2; i < 10; i++)
                {
                    var level = new Level
                    {
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = level1.NumberingFormat.Val },
                        LevelIndex = i - 1,
                        LevelText = new LevelText() { Val = $"%{i}." }
                    };

                    if (cascading)
                    {
                        // if we're cascading, that means we don't want any identation 
                        // + our leveltext should contain the previous levels as well
                        var lvlText = new StringBuilder();

                        for (int lvlIndex = 1; lvlIndex <= i; lvlIndex++)
                            lvlText.AppendFormat("%{0}.", lvlIndex);

                        level.LevelText = new LevelText() { Val = lvlText.ToString() };
                    }
                    else
                    {
                        level.PreviousParagraphProperties = new PreviousParagraphProperties
                        {
                            Indentation = new Indentation() { Left = (720 * i).ToString(CultureInfo.InvariantCulture), Hanging = "360" }
                        };
                    }

                    absNumMultilevel.Append(level);
                }
            }
        }

        #endregion

        //____________________________________________________________________
        //
        // Properties Implementation
        #region Properties

        /// <summary> Gets the depth level of the current list instance. </summary>
        public int LevelIndex => levelDepth;

        /// <summary>  </summary>
        public string[] CurrentListClasses => listHtmlElementClasses.Peek();

        /// <summary> Gets the ID of the current list instance. </summary>
        internal int InstanceId => numInstances.Peek().Key;

        internal IEnumerable<AbstractNum> AbstractNums
            => mainPart?.NumberingDefinitionsPart?.Numbering?.Elements<AbstractNum>()
            ?? Enumerable.Empty<AbstractNum>();

        /// <summary>  </summary>
        public string[] OrderedTypes => new[] {
            OrderingTypeDecimal,
            OrderingTypeUpperAlpha,
            OrderingTypeLowerAlpha,
            OrderingTypeUpperRoman,
            OrderingTypeLowerRoman,
            HEADING_NUMBERING_NAME
        };

        #endregion
    }
}