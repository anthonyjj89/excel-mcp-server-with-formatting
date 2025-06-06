package tools

import (
	"context"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
  "github.com/negokaz/excel-mcp-server/internal/excel"
)

type ReadSheetNameArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
}

var readSheetNameArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Required(),
})

func AddReadSheetNamesTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("read_sheet_names",
		mcp.WithDescription("List all sheet names in an Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
	), handleReadSheetNames)
}

func handleReadSheetNames(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ReadSheetNameArguments{}
	issues := readSheetNameArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return readSheetNames(args.FileAbsolutePath)
}

func readSheetNames(fileAbsolutePath string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
  defer release()
	if err != nil {
		return nil, err
	}

	sheetList, err := workbook.GetSheetNames()
  if err != nil {
    return nil, err
  }
	var sheetNames []mcp.Content
	for _, name := range sheetList {
		sheetNames = append(sheetNames, mcp.NewTextContent(name))
	}

	return &mcp.CallToolResult{
		Content: sheetNames,
	}, nil
}
