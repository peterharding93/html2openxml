﻿/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Defines the styles to apply on OpenXml elements.
	/// </summary>
	public sealed class HtmlDocumentStyle
	{
		/// <summary>
		/// Occurs when a Style is missing in the MainDocumentPart but will be used during the conversion process.
		/// </summary>
		public event EventHandler<StyleEventArgs> StyleMissing;

		/// <summary>
		/// Contains the default styles for new OpenXML elements
		/// </summary>
		public DefaultStyles DefaultStyles { get { return this.defaultStyles; } }

		private DefaultStyles defaultStyles = new DefaultStyles();
		private RunStyleCollection runStyle;
		private TableStyleCollection tableStyle;
		private ParagraphStyleCollection paraStyle;
        	private NumberingListStyleCollection listStyle;
		private OpenXmlDocumentStyleCollection knownStyles;
		private MainDocumentPart mainPart;


		internal HtmlDocumentStyle(MainDocumentPart mainPart)
		{
			PrepareStyles(mainPart);
			tableStyle = new TableStyleCollection(this);
			runStyle = new RunStyleCollection(this);
			paraStyle = new ParagraphStyleCollection(this);
            		this.QuoteCharacters = QuoteChars.IE;
			this.mainPart = mainPart;
		}

		//____________________________________________________________________
		//

		#region PrepareStyles

        /// <summary>
        /// Preload the styles in the document to match localized style name.
        /// </summary>
        internal void PrepareStyles(MainDocumentPart mainPart)
        {
            knownStyles = new OpenXmlDocumentStyleCollection();
            if (mainPart.StyleDefinitionsPart == null) return;

            Styles styles = mainPart.StyleDefinitionsPart.Styles;

            foreach (var s in styles.Elements<Style>())
            {

#if false

				/* This version ensures that the style is avaialble both by NAME and by Id.
				 * This was the originbal version
				 * That runs the risk of the dictionary key not being unique.
				 */
                StyleName n = s.StyleName;
                if (n != null)
                {
                    String name = n.Val.Value;
                    if (name != s.StyleId) knownStyles.Add(name, s);
                }
#endif

#if false
				/* This code changes the Id of existing styles.
				 * This breaks documents with existing content using the existing styles.
				 * This code seems to do two things:
				 *		1) changes StyleId to match the Name 
				 *		2) generates unique StyleId (by appending numeric suffix) if there is a StyleId conflict.
				 * I'm guessing that this was a workaround for other pieces of code which assumed Id is the name.
				 * Correct solution would be to fix any modules that make that assumption.
				 * Also a bit strange because the unique generate Id would no longer match the name.
				 * 
				 * Suggested behaviour (and test case):
				 *    Merely constructing HtmlConverter(document) should not mutate the document.				 
				 */

				StyleName n = s.StyleName;
                string originalIdName = s.StyleId;
                var id = 1;

                if (n != null)
                {
                    string name = n.Val.Value;
                    if (name != originalIdName)
                    {
                        originalIdName = name;
                    }
                }

                s.StyleId = originalIdName;

                while (knownStyles.ContainsKey(s.StyleId))
                {
                    id++;
                    s.StyleId = originalIdName + id.ToString("00");
                }
#endif
                knownStyles.Add(s.StyleId, s);
            }

#if true
			/* This version will only add lookup by Name if there isn't a naming conflict.
			 * Do this as a second pass.
			 * If there is a naming conflict, then oh well.
			 */
			foreach (var s in styles.Elements<Style>())
			{
				if (s.StyleName is null)
					continue;
                
				String name = s.StyleName.Val.Value;

                if (!knownStyles.ContainsKey(name))
				{
					knownStyles.Add(name, s);
				}
                
            }

#endif


            }

#endregion

		#region GetStyle

        /// <summary>
        /// Helper method to obtain the StyleId of a named style (invariant or localized name), of type Paragraph, case sensitive
        /// </summary>
        /// <param name="name">The name of the style to look for.</param>
        public String GetStyle(string name) => GetStyle(name, StyleValues.Paragraph);
		
        /// <summary>
        /// Helper method to obtain the StyleId of a named style (invariant or localized name).
        /// </summary>
        /// <param name="name">The name of the style to look for.</param>
        /// <param name="styleType">True to obtain the character version of the given style.</param>
        /// <param name="ignoreCase">Indicate whether the search should be performed with the case-sensitive flag or not.</param>
        /// <returns>If not found, returns the given name argument.</returns>
        public String GetStyle(string name, StyleValues styleType, bool ignoreCase = false)
		{
			Style style;

			// OpenXml is case-sensitive but CSS is not.
			// We will try to find the styles another time with case-insensitive:
			if (ignoreCase)
			{
				if (!knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
				{
					if (StyleMissing != null)
					{
						StyleMissing(this, new StyleEventArgs(name, mainPart, styleType));
						if (knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
							return style.StyleId;
					}
					return null; // null means we ignore this style (css class)
				}

				return style.StyleId;
			}
			else
			{
				if (!knownStyles.TryGetValue(name, out style))
				{
					if (!EnsureKnownStyle(name, out style))
					{
						StyleMissing?.Invoke(this, new StyleEventArgs(name, mainPart, styleType));
						return name;
					}
				}

				if (styleType == StyleValues.Character && !style.Type.Equals<StyleValues>(StyleValues.Character))
				{
					LinkedStyle linkStyle = style.GetFirstChild<LinkedStyle>();
					if (linkStyle != null) return linkStyle.Val;
				}
				return style.StyleId;
			}
		}

		#endregion

		#region DoesStyleExists

		/// <summary>
		/// Gets whether the given style exists in the document.
		/// </summary>
		public bool DoesStyleExists(string name)
		{
			return knownStyles.ContainsKey(name);
		}

		#endregion

		#region AddStyle

		/// <summary>
		/// Add a new style inside the document and refresh the style cache.
		/// </summary>
		internal void AddStyle(String name, Style style)
		{
			knownStyles[name] = style;
			if (mainPart.StyleDefinitionsPart == null)
				mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
			mainPart.StyleDefinitionsPart.Styles.Append(style);
		}

		#endregion

		#region EnsureKnownStyle

        /// <summary>
        /// Try to insert the style in the document if it is a known style.
        /// </summary>
        private bool EnsureKnownStyle(string styleName, out Style style)
        {
			style = null;
			string xml = PredefinedStyles.GetOuterXml(styleName);
			if (xml == null) return false;
			this.AddStyle(styleName, style = new Style(xml));
			return true;
        }

		#endregion

		//____________________________________________________________________
		//

		internal RunStyleCollection Runs
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return runStyle; }
		}
		internal TableStyleCollection Tables
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return tableStyle; }
		}
		internal ParagraphStyleCollection Paragraph
		{
			[System.Diagnostics.DebuggerHidden()]
			get { return paraStyle; }
		}
        internal NumberingListStyleCollection NumberingList
        {
			// use lazy loading to avoid injecting NumberListDefinition if not required
            [System.Diagnostics.DebuggerHidden()]
            get { return listStyle ?? (listStyle = new NumberingListStyleCollection(mainPart)); }
        }

		//____________________________________________________________________
		//

        /// <summary>
        /// Gets or sets the beginning and ending characters used in the &lt;q&gt; tag.
        /// </summary>
        public QuoteChars QuoteCharacters { get; set; }
	}
}
