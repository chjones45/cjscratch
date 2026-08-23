-- Force a specific DOCX table style for all Markdown tables.
-- Set style name with either:
--   -M docx-table-style=YourStyleName
--   -M docx_table_style=YourStyleName

local docx_table_style = "NetApp Table 1"

function Meta(meta)
  local configured = nil

  if meta["docx-table-style"] then
    configured = pandoc.utils.stringify(meta["docx-table-style"])
  elseif meta["docx_table_style"] then
    configured = pandoc.utils.stringify(meta["docx_table_style"])
  end

  if configured and configured ~= "" then
    docx_table_style = configured
  end

  return meta
end

function Table(tbl)
  tbl.attr = tbl.attr or pandoc.Attr()
  tbl.attr.attributes = tbl.attr.attributes or {}
  tbl.attr.attributes["custom-style"] = docx_table_style
  tbl.attr.attributes["table-style"] = docx_table_style
  return tbl
end
