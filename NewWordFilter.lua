---
--- Various extras for converting Markdown to Docx
---

--- Tables
--		Support creating auto-numbered captions
--			also create a bookmark for each one
--	Figures
--		Support creating auto-numbered captions
--			also create a bookmark for each one
--	Links
--		Create a link to any pre-defined bookmark
--	Definition Lists
--		Automatically number and config per ISO
--[[
Author: Leonard Rosenthol
Copyright (c) 2020, Adobe
]]


--------------
-- load some libraries
local List = require 'pandoc.List'
local stringify = require("pandoc.utils").stringify


--------------
-- local vars
local TABLE_COUNT = 0
local FIGURE_COUNT = 0
local BM_COUNT = 0

-- local OOXML code
local RAW_TOC = [[
<w:sdt>
    <w:sdtContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
            <w:r>
                <w:fldChar w:fldCharType="begin" w:dirty="true" />
                <w:instrText xml:space="preserve">TOC \o "1-3" \h \z \u</w:instrText>
                <w:fldChar w:fldCharType="separate" />
                <w:fldChar w:fldCharType="end" />
            </w:r>
        </w:p>
    </w:sdtContent>
</w:sdt>
]]
local RAW_PAGEBREAK = "<w:p><w:r><w:br w:type=\"page\" /></w:r></w:p>"
local RAW_BOOKMARK = "<w:bookmarkStart w:id=\"%d\" w:name=\"%s\"/><w:r><w:t>%s</w:t></w:r><w:bookmarkEnd w:id=\"%d\"/>"
local RAW_LINK = "<w:hyperlink w:anchor=\"%s\"><w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>%s</w:t></w:r></w:hyperlink>"

local RAW_TERM_OPEN = '<w:pPr><w:pStyle w:val="Term(s)"/></w:pPr><w:r><w:t>'
local RAW_TERM_CLOSE = '</w:t></w:r>'
local RAW_LINEBREAK = '<w:br />'

--------------
-- private methods

local function debug(string)
    io.stderr:write("[Debug] " .. string .. "\n")
end

local function docx_inline(text)
    return pandoc.RawInline("openxml", text)
end

local function docx_block(text)
    return pandoc.RawBlock("openxml", text)
end

local function make_bookmark(bmname, content)
	BM_COUNT = BM_COUNT + 1
	bmtext = string.format(RAW_BOOKMARK, BM_COUNT, bmname, pandoc.utils.stringify(content), BM_COUNT)
	return {docx_inline(bmtext)}
end

local function make_link(lname, content)
	linktext = string.format(RAW_LINK, lname, pandoc.utils.stringify(content))
	return {docx_inline(linktext)}
end

--- Parses a mock "Table Attr".
-- We use the Attr of an empty Span as if it were Table Attr.
-- This function extracts what is needed to build a short-caption.
-- @tparam Attr attr : The Attr of the property Span in the table caption
-- @treturn ?string : The identifier
-- @treturn ?string : The "short-caption" property, if present.
-- @treturn bool : Whether ".unlisted" appeared in the classes
local function parse_table_attrs(attr)
	-- Find label
	local label = nil
	if attr.identifier and (#attr.identifier > 0) then
	  label = attr.identifier
	end
  
	-- Look for ".unlisted" in classes
	local unlisted = false
	if attr.classes:includes("unlisted") then
	  unlisted = true
	end
  
	-- If not unlisted, then find the property short-caption.
	local short_caption = nil
	if not unlisted then
	  if (attr.attributes["short-caption"]) and
		 (#attr.attributes["short-caption"] > 0) then
		short_caption = attr.attributes['short-caption']
	  end
	end
  
	return label, short_caption, unlisted
  end
  
--------------
-- methods for processing specific elements

local function do_raw_block(el)
    if el.text == "\\toc" then
        if FORMAT == "docx" then
            debug("inserting Table of Contents")
            el.text = RAW_TOC
            el.format = "openxml"
            local para = pandoc.Para( pandoc.Str 'Table of Contents' )
            local div = pandoc.Div({ para, el })
            div["attr"]["attributes"]["custom-style"] = "TOC Heading"
            return div
        else
            --debug("\\toc, not docx")
            return {}
        end
    elseif el.text == "\\newpage" then
        if FORMAT == "docx" then
            debug("inserting a Pagebreak")
            el.text = RAW_PAGEBREAK
            el.format = "openxml"
            return el
        else
            --debug("\\newpage, not docx")
            return {}
        end
    end
end

local function do_table(tbl)
	local caption
	if PANDOC_VERSION >= {2,10} then
	  caption = pandoc.List(tbl.caption.long)
	else
	  caption = tbl.caption
	end

	-- Escape if there is no caption present.
	if not caption or #caption == 0 then
		debug("No Caption")
		return nil
	else 
		-- debug("Has Caption: " .. pandoc.utils.stringify(caption))
	end

	local label
	if PANDOC_VERSION >= {2,10} then
		_, _, captionTxt, label = string.find(pandoc.utils.stringify(caption), "(.+){#(.+)}")
	else
		-- Try find the properties block
		local is_properties_span = function (inl)
			return (inl.t) and (inl.t == "Span")   						-- is span
						and (inl.content) and (#inl.content == 0)  	-- is empty span	 
		end
		local propspan, idx = caption:find_if(is_properties_span)

		-- If we couldn't find properties, escape.
		if not propspan then
			debug("No propspan")
			return nil
		end

		-- Otherwise, parse it all
		label, short_caption, unlisted = parse_table_attrs(propspan.attr)

		caption[idx] = nil
	end
	-- debug("Label: "..label)
	
	-- if we get here, then we have a table that needs a number
	-- Put label back into caption, with new name & number
	if label then
		TABLE_COUNT = TABLE_COUNT + 1
		local tableLbl = string.format("Table %d: ", TABLE_COUNT)
		if PANDOC_VERSION >= {2,10} then
			caption = { pandoc.Str(tableLbl), pandoc.Str(captionTxt) }
			captionTxt = pandoc.utils.stringify(caption)
			-- debug("new Caption: " .. captionTxt)
		else
			caption:insert( 1, pandoc.Str(tableLbl) )
		end
		caption = make_bookmark(label, caption)
		short_caption = caption
	end

	-- set new caption
	if PANDOC_VERSION >= {2,10} then
	  	tbl.caption.long = pandoc.Para( captionTxt )
		tbl.caption.short = short_caption

		debug("Table Caption: " .. pandoc.utils.stringify(tbl.caption.long))
		-- debug("Short Caption: "..pandoc.utils.stringify(tbl.caption.short))
	else
	  tbl.caption = caption
	  -- debug("Caption: "..pandoc.utils.stringify(tbl.caption))
	end

	-- Place new table
	local result = List:new{}

	-- this block is required as the current Docx exporter doesn't handle the new
	-- style of captions correctly!
	if PANDOC_VERSION >= {2,10} then
		local caption_div = pandoc.Div({})
		caption_div["attr"]["attributes"]["custom-style"] = "Table Caption"
		caption_div.content = { pandoc.Para(caption) }
		result:extend{ caption_div }
	end

	result:extend {tbl}
	return result
end  

local function do_span(sp)
    local spanType = sp.classes[1]

	if spanType == "link" then
		-- pull the name of the bookmark out ("foo!name")
		content = pandoc.utils.stringify(sp.content)
		_, _, content, lname = string.find(content, "(.+)!(.+)")

		-- compute it and set it as the new span
		return make_link(lname, content)
	elseif spanType == "image" then
		debug("image span")
		return nil
	end
end

local function do_para(elem)
	-- debug("para: " .. pandoc.utils.stringify(elem))

	if #elem.content == 1 and elem.content[1].tag == "Image" then
		debug("Para having one Image element found")
		image = elem.content[1]
		--debug(stringify(image.src))
		local caption_div = pandoc.Div({})
		local image_div = pandoc.Div({})
		-- caption_div["attr"]["attributes"]["custom-style"] = stringify(meta["caption"])
		-- image_div["attr"]["attributes"]["custom-style"] = stringify(meta["anchor"])

		if stringify(image.caption) ~= "" then
			debug(stringify(image.caption))

			content = pandoc.utils.stringify(image.caption)
			_, _, content, bmname = string.find(content, "(.+)#(.+)")

			-- just in case...
			if content and bmname then
				debug("Caption: " .. stringify(content))
				debug("Bookmark: " .. stringify(bmname))

				FIGURE_COUNT = FIGURE_COUNT + 1
				local figLbl = string.format("Figure %d: ", FIGURE_COUNT)
				image.caption = { pandoc.Str(figLbl), pandoc.Str(stringify(content)) }
				debug("new Caption: " .. stringify(image.caption))
				image.caption = make_bookmark(bmname, image.caption)

				caption_div.content = { pandoc.Para(image.caption) }
				image.caption = {}
				image.title = ""
			end
		end
		image_div.content = { pandoc.Para(image) }
		return { image_div, caption_div }
	end
end

local function do_def_list(dl)
	local outList = {}

	for i, item in ipairs(dl.content) do
		-- debug(string.format("Found item %d - %s | %s", i, pandoc.utils.stringify(item[1]), item[2]))

		local termNum = string.format("3.%d", i)
		local outTerm = docx_inline( RAW_TERM_OPEN .. 
								termNum .. 
								RAW_LINEBREAK ..
								pandoc.utils.stringify(item[1]) ..
								RAW_TERM_CLOSE)
		local outDef = item[2];
		
		table.insert(outList, {outTerm, outDef})
	end    

	return pandoc.DefinitionList(outList)
end

local function do_meta(mt)
	-- for each meta listed here, remove it from the output
	local meta = {
		author = "author-meta",
		date = "date-meta",
		subtitle = "subtitle-meta",
		title = "title-meta",
	}

	for k, v in pairs(meta) do
		--debug(k .. ": " .. v)
		if mt[k] ~= nil then
			mt[v] = stringify(mt[k])
			--debug(stringify(mt[k]))
			mt[k] = nil
			debug(string.format("metadata '%s' has found and removed", k))
		end
	end
	--pretty.dump(mt)
	return mt
end

------------------------
-- don't do anything unless we target docx
if FORMAT ~= "docx" then
	return {}
  end
  
-- this is the "main" Pandoc routine that connects the parts of the doc->methods
return {
	{
	  Meta = do_meta,
	  Table = do_table,
	  RawBlock = do_raw_block,
	  DefinitionList = do_def_list,
	  Span = do_span,
	  Para = do_para
	}
  }