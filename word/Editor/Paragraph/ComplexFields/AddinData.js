/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function(window)
{
	/**
	 * Class for working with adding field data
	 * @constructor
	 */
	function CAddinFieldData()
	{
		this.FieldId = undefined;
		this.Value   = undefined;
		this.Content = undefined;
		// Hearth (fork): the result as formatted paragraphs (SetParagraphs).
		this.Paragraphs = undefined;
	}
	CAddinFieldData.prototype.SetFieldId = function(fieldId)
	{
		this.FieldId = "" + fieldId;
	};
	CAddinFieldData.prototype.GetFieldId = function()
	{
		return this.FieldId;
	};
	CAddinFieldData.prototype.SetValue = function(value)
	{
		this.Value = "" + value;
	};
	CAddinFieldData.prototype.GetValue = function()
	{
		return this.Value;
	};
	CAddinFieldData.prototype.SetContent = function(content)
	{
		this.Content = "" + content;
	};
	CAddinFieldData.prototype.GetContent = function()
	{
		return this.Content;
	};
	/**
	 * Hearth (fork): the field's result as paragraphs of formatted runs, in
	 * place of the plain Content: [{"Runs": [{"Text", "Bold", "Italic",
	 * "Underline", "SmallCaps", "VertAlign": "superscript" | "subscript" |
	 * "baseline"}], "IndLeft", "IndFirstLine" (twips)}]. A citation's
	 * superscript number and a bibliography's entries (one paragraph each,
	 * with italic journal names and a hanging indent) are written in one
	 * step, so neither is lost or broken when a field is updated.
	 * @param {Array} paragraphs
	 */
	CAddinFieldData.prototype.SetParagraphs = function(paragraphs)
	{
		this.Paragraphs = Array.isArray(paragraphs) ? paragraphs : undefined;
	};
	CAddinFieldData.prototype.GetParagraphs = function()
	{
		return this.Paragraphs;
	};
	CAddinFieldData.prototype.ToJson = function()
	{
		return {
			"FieldId" : this.FieldId,
			"Value"   : this.Value,
			"Content" : this.Content
		};
	};
	CAddinFieldData.FromJson = function(obj)
	{
		let newData = new CAddinFieldData();
		if (!obj)
			return newData;
		
		if (undefined !== obj.FieldId)
			newData.SetFieldId(obj.FieldId);
		else if (undefined !== obj["FieldId"])
			newData.SetFieldId(obj["FieldId"]);
		
		if (undefined !== obj.Value)
			newData.SetValue(obj.Value);
		else if (undefined !== obj["Value"])
			newData.SetValue(obj["Value"]);
		
		if (undefined !== obj.Content)
			newData.SetContent(obj.Content);
		else if (undefined !== obj["Content"])
			newData.SetContent(obj["Content"]);
		
		if (undefined !== obj["Paragraphs"])
			newData.SetParagraphs(obj["Paragraphs"]);
		
		return newData;
	};
	CAddinFieldData.FromField = function(field)
	{
		if (!field
			|| !(field instanceof AscWord.CComplexField)
			|| !field.IsAddin()
			|| null === field.GetFieldId())
			return null;
		
		let data = new CAddinFieldData();
		data.SetFieldId(field.GetFieldId());
		data.SetValue(field.GetInstruction().GetValue());
		data.SetContent(field.GetFieldValueText());
		return data;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CAddinFieldData = CAddinFieldData;
	
})(window);

