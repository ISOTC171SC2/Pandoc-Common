--[[

----------------------------------------------------------------------------]]

-- shorthand for setting raw OOXML
function docx(text)
    return pandoc.RawInline("openxml", text)
end

function docx_block(text)
    return pandoc.RawBlock("openxml", text)
end

-- this is the set of DocX-specific info we will inject
local DOCX_TEXT = {}
DOCX_TEXT.block_comment = {}
DOCX_TEXT.block_comment.Open = docx('')
DOCX_TEXT.block_comment.Close = docx('')
DOCX_TEXT.block_box = {}
DOCX_TEXT.block_box.Open = docx('<w:pPr><w:pStyle w:val="Box"/></w:pPr><w:r><w:t>')
DOCX_TEXT.block_box.Close = docx('</w:t></w:r>')
DOCX_TEXT.block_center = {}
DOCX_TEXT.block_center.Open = docx('')
DOCX_TEXT.block_center.Close = docx('')
DOCX_TEXT.block_toc = {}
DOCX_TEXT.block_toc.Open = docx('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents" /><w:docPartUnique /></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOCHeading" /></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true" /><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u')
DOCX_TEXT.block_toc.Close = docx('</w:instrText><w:fldChar w:fldCharType="separate" /><w:fldChar w:fldCharType="end" /></w:r></w:p></w:sdtContent></w:sdt>')
DOCX_TEXT.block_ednote = {}
DOCX_TEXT.block_ednote.Open = docx('<w:pPr><w:pStyle w:val="Editors Note"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:pPr><w:r><w:t>')
DOCX_TEXT.block_ednote.Close = docx('</w:t></w:r>')
DOCX_TEXT.comment = {}
DOCX_TEXT.comment.Open = docx('<w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>')
DOCX_TEXT.comment.Close = docx('</w:t>')
DOCX_TEXT.highlight = {}
DOCX_TEXT.highlight.Open = docx('<w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>')
DOCX_TEXT.highlight.Close = docx('</w:t>')
DOCX_TEXT.code = {}
DOCX_TEXT.code.Open = docx('<w:rPr><w:rStyle w:val="InlineCode"/></w:rPr><w:t>')
DOCX_TEXT.code.Close = docx('</w:t>')
DOCX_TEXT.note = {}
DOCX_TEXT.note.Open = docx('<w:rPr><w:rStyle w:val="InlineNote"/></w:rPr><w:t>')
DOCX_TEXT.note.Close = docx('</w:t>')
DOCX_TEXT.margin = {}
DOCX_TEXT.margin.Open = docx('')
DOCX_TEXT.margin.Close = docx('')
DOCX_TEXT.fixme = {}
DOCX_TEXT.fixme.Open = docx('<w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>')
DOCX_TEXT.fixme.Close = docx('</w:t>')
DOCX_TEXT.noindent = docx('')
DOCX_TEXT.i = {}
DOCX_TEXT.i.Open = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> XE "</w:instrText></w:r><w:r><w:instrText>'
DOCX_TEXT.i.Close = '</w:instrText></w:r><w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
DOCX_TEXT.link = {}
DOCX_TEXT.link.Open = ''
DOCX_TEXT.link.Close = ''
DOCX_TEXT.bookmark = {}
DOCX_TEXT.bookmark.Open = ''
DOCX_TEXT.bookmark.Close = ''
DOCX_TEXT.rp = {}
DOCX_TEXT.rp.Open = ''
DOCX_TEXT.rp.Close = ''
DOCX_TEXT.term = {}
DOCX_TEXT.term.Open = '<w:pPr><w:pStyle w:val="Term(s)"/></w:pPr><w:r><w:t>'
DOCX_TEXT.term.Close = '</w:t></w:r>'
DOCX_TEXT.linebreak = '<w:br />'

-- Used to store YAML variables
local YAML_VARS = {}

-- Used to count words in text, abstract, and footnotes
local WORD_COUNT = 0
local ABSTRACT_COUNT = 0
local NOTE_COUNT = 0
local YAML_WORDS = 0 -- Used for counting # words in YAML values (to be subtracted)
local BOOKMARK_COUNT = 0

function isWord(text)
    -- Returns true/false if text contains word characters (not just punctuation)
    return text:match("%P")
end

function loadYAML(meta)
    -- Record metadata for later use, and count words.
    for key, value in pairs(meta) do
        YAML_VARS[key] = value
        if type(value) ~= "boolean" then
            -- count words in YAML header, keeping track of those in abstract.
            if value.t == "MetaBlocks" then
                for _, block in pairs(value) do
                    pandoc.walk_block(block, {
                        Str = function(string)
                            if isWord(string.text) then
                                YAML_WORDS = YAML_WORDS + 1
                                if key == "abstract" then
                                    ABSTRACT_COUNT = ABSTRACT_COUNT + 1
                                end
                            end
                            return
                        end})
                end
            elseif value.t == "MetaList" then
                for _, item in pairs(value) do
                    for _, inline in pairs(item) do
                        if inline.t == "Str" and isWord(inline.text) then
                            YAML_WORDS = YAML_WORDS + 1
                        end
                    end
                end
            elseif value.t == "MetaInlines" then
                for _, inline in pairs(value) do
                    if inline.t == "Str" and isWord(inline.text) then
                        YAML_WORDS = YAML_WORDS + 1
                        if key == "abstract" then
                            ABSTRACT_COUNT = ABSTRACT_COUNT + 1
                        -- elseif key == "tempdir" then
                        --     -- Set IMAGE_PATH from metadata
                        --     IMAGE_PATH = pandoc.utils.stringify(value.c) .. '/Figures/'
                        end
                    end
                end
            end
        end
    end
end

function setMeta(meta)
    -- Revise document metadata as appropriate; print detailed wordcount.
    if FORMAT == "markdown" then  -- Don't change anything if translating to .md
        return
	end

	-- set the date to the current one
    meta["date"] = os.date("%B %e, %Y")

    -- print out info about the document
    print(string.format("Words: %d │ Abstract: %d │ Notes: %d │ Body: %d",
          WORD_COUNT - YAML_WORDS + ABSTRACT_COUNT, ABSTRACT_COUNT, NOTE_COUNT,
          WORD_COUNT - NOTE_COUNT - YAML_WORDS))
    return meta
end

function handleDL(dl)
    -- we only care about DocX in this filter
    if FORMAT == "docx" then
        local outList = {}

        for i, item in ipairs(dl.content) do
            -- print(string.format("Found item %d - %s | %s", i, pandoc.utils.stringify(item[1]), item[2]))

            local termNum = string.format("3.%d", i)
            local outTerm = docx( DOCX_TEXT.term.Open .. 
                                    termNum .. 
                                    DOCX_TEXT.linebreak ..
                                    pandoc.utils.stringify(item[1]) ..
                                    DOCX_TEXT.term.Close)
            local outDef = item[2];
            
            table.insert(outList, {outTerm, outDef})
        end    

        return pandoc.DefinitionList(outList)
    end
end

function isSpecialBlock(text)
    return  text == 'comment' or text == 'box' or text == 'center' or 
            text == 'toc' or text == 'ednote'
end

function handleBlocks(block)
    if FORMAT == "markdown" then  -- Don't change anything if translating to .md
        return
    elseif isSpecialBlock(block.classes[1]) then
		-- we only care about DocX in this filter
        if FORMAT == "docx" then
            return
                {pandoc.Plain({DOCX_TEXT["block_" .. block.classes[1]].Open})} ..
                block.content ..
                {pandoc.Plain({DOCX_TEXT["block_" .. block.classes[1]].Close})}
            -- insert in front
            -- table.insert( block.content, 1, pandoc.Plain({DOCX_TEXT["block_" .. block.classes[1]].Open}) )
            -- insert in back
            -- table.insert( block.content, pandoc.Plain({DOCX_TEXT["block_" .. block.classes[1]].Close}) )
            -- return block
        end
    elseif block.classes[1] == nil then
        -- print(string.format("c: '%s'\t c1: '%s'", block.content, block.content[1]))
        -- print(string.format("attr '%s'", block.attributes.number))
        
        if block.classes[2] ~= nil then
            foo = block.classes[2][0];
            print(string.format("foo: %s", foo))
        end
    end
end

function isSpecialSpan(text)
    return text == 'comment' or text == 'margin' or 
            text == 'fixme' or text == 'highlight' or
            text == 'code' or text == 'note'
end

function handleInlines(span)
    if FORMAT == "markdown" then  -- Don't change anything if translating to .md
        return
    end
    local spanType = span.classes[1]
    if isSpecialSpan(spanType) then
 		-- we only care about DocX in this filter
		 if FORMAT == "docx" then
            return
                {DOCX_TEXT[spanType].Open} ..
                span.content ..
                {DOCX_TEXT[spanType].Close}
        end
    elseif spanType == "smcaps" then
        return pandoc.SmallCaps(span.content)
    elseif spanType == "i" then
		-- Process indexing ...
		-- we only care about DocX in this filter
        if FORMAT == 'docx' then
            print(span.content)
            local indexItem = pandoc.utils.stringify(span.content)
            indexItem = string.gsub(indexItem, '!', ':')  -- Subheadings
            -- Need to do other things here to identify ranges of pages,
            -- "See" and "See also" cross references, etc.
            return docx(
                DOCX_TEXT.i.Open ..
                indexItem ..
                DOCX_TEXT.i.Close)
        else
            return {}
        end
    elseif spanType == "l" or spanType == "r" or spanType == "rp" then
        -- Process cross-references ...
        content = pandoc.utils.stringify(span.content)

		-- we only care about DocX in this filter
		if FORMAT == "docx" then
            return {docx(
                DOCX_TEXT[spanType].Open ..
                content ..
                DOCX_TEXT[spanType].Close)}
        end
    elseif spanType == "bookmark" then
		-- we only care about DocX in this filter
        if FORMAT == "docx" then
            BOOKMARK_COUNT = BOOKMARK_COUNT + 1

            -- pull the name of the BM out ("foo!name")
            content = pandoc.utils.stringify(span.content)
            _, _, content, bmname = string.find(content, "(.+)!(.+)")

            -- fill it all in...
            bmtext = string.format("<w:bookmarkStart w:id=\"%d\" w:name=\"%s\"/><w:r><w:t>%s</w:t></w:r><w:bookmarkEnd w:id=\"%d\"/>",
                                    BOOKMARK_COUNT, bmname, content, BOOKMARK_COUNT)
            -- print(bmtext)
            return {docx(bmtext)}
        end
    elseif spanType == "link" then
		-- we only care about DocX in this filter
        if FORMAT == "docx" then
            -- pull the name of the BM out ("foo!name")
            content = pandoc.utils.stringify(span.content)
            _, _, content, bmname = string.find(content, "(.+)!(.+)")

            -- fill it all in...
            linktext = string.format("<w:hyperlink w:anchor=\"%s\"><w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>%s</w:t></w:r></w:hyperlink>",
                                    bmname, content)
            -- print(linktext)
            return {docx(linktext)}
        end
    end
end

--- configs
local pagebreak = {
    epub = '<p style="page-break-after: always;"> </p>',
    html = '<div style="page-break-after: always;"></div>',
    latex = '\\newpage{}',
    ooxml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    odt = '<text:p text:style-name="Pagebreak"/>'
  }

--- Return a block element causing a page break in the given format.
local function newpage(format)
	if format == 'docx' then
		return pandoc.RawBlock('openxml', pagebreak.ooxml)
	elseif format:match 'latex' then
		return pandoc.RawBlock('tex', pagebreak.latex)
	elseif format:match 'odt' then
		return pandoc.RawBlock('opendocument', pagebreak.odt)
	elseif format:match 'html.*' then
		return pandoc.RawBlock('html', pagebreak.html)
	elseif format:match 'epub' then
		return pandoc.RawBlock('html', pagebreak.epub)
	else
		-- fall back to insert a form feed character
		return pandoc.Para{pandoc.Str '\f'}
	end
end

local function is_newpage_command(command)
	return command:match '^\\newpage%{?%}?$'
		or command:match '^\\pagebreak%{?%}?$'
end

-- Filter function called on each RawBlock element.
function RawBlock (el)
	-- Don't do anything if the output is TeX
	if FORMAT:match 'tex$' then
		return nil
	end
	-- check that the block is TeX or LaTeX and contains only
	-- \newpage or \pagebreak.
	if el.format:match 'tex' and is_newpage_command(el.text) then
		-- use format-specific pagebreak marker. FORMAT is set by pandoc to
		-- the targeted output format.
		return newpage(FORMAT)
	end
	-- otherwise, leave the block unchanged
	return nil
end

-- Turning paragraphs which contain nothing but a form feed
-- characters into line breaks.
function NLPara (el)
	if #el.content == 1 and el.content[1].text == '\f' then
		return newpage(FORMAT)
	end
end
  
function handleNotes(note)
    return pandoc.walk_inline(note, {
        Str = function(string)
            if isWord(string.text) then
                NOTE_COUNT = NOTE_COUNT + 1
            end
            return
        end})
end


function handleStrings(string)
    if isWord(string.text) then  -- If string contains non-punctuation chars
        WORD_COUNT = WORD_COUNT + 1  -- ... count it.
    end
    return
end


-- Order matters here!
local COMMENT_FILTER = {
    {Meta = loadYAML},             -- This comes first to read metadata values
--    {Para = handleTransclusion},  -- Transclusion before other filters
--    {Para = handleNoIndent},      -- Non-indented paragraphs (after transclusion)
--    {CodeBlock = handleCode},     -- Convert TikZ images (before Image)
    {DefinitionList = handleDL},    -- definition lists
    {Div = handleBlocks},         -- Comment blocks (before inlines)
    {RawBlock = RawBlock, Para = NLPara},  -- Handle forced page breaks
--    {Image = handleImages},       -- Images (so captions get inline filters)
--    {Math = handleMacros},        -- Replace macros from YAML data
    {Span = handleInlines},       -- Comment and cross-ref inlines
    {Note = handleNotes},         -- Count words
    {Str = handleStrings},        -- Count words
    {Meta = setMeta}              -- This comes last to rewrite YAML
}

return COMMENT_FILTER
