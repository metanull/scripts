[CmdletBinding(DefaultParameterSetName='GetAll')]
param(
    [Parameter(ParameterSetName = 'GetConstant')]
    [string]$ConstantName,

    [Parameter(ParameterSetName = 'GetAll')]
    [switch]$All,

    [Parameter(ParameterSetName = 'ListConstants')]
    [switch]$ListAvailable
)

$WordConstants = @{
    WD_LINE_STYLE_NONE = 0
    WD_LINE_STYLE_SINGLE = 1
    WD_SECTION_BREAK_NEXT_PAGE = 2
    WD_LINE_WIDTH_075PT = 6
    WD_AUTOFIT_FIXED = 0
    WD_PREFERRED_WIDTH_PERCENT = 1
    # Additional commonly used Word constants
    WD_PAGE_BREAK = 1
    WD_LINE_STYLE_DOUBLE = 7
    WD_LINE_WIDTH_150PT = 12
    WD_LINE_WIDTH_225PT = 18
    WD_LINE_WIDTH_300PT = 24
    WD_BORDER_TOP = 1
    WD_BORDER_LEFT = 2
    WD_BORDER_BOTTOM = 3
    WD_BORDER_RIGHT = 4
    WD_BORDER_HORIZONTAL = 5
    WD_BORDER_VERTICAL = 6
    WD_ROW_HEIGHT_AUTO = 0
    WD_ROW_HEIGHT_AT_LEAST = 1
    WD_ROW_HEIGHT_EXACTLY = 2
}

switch ($PSCmdlet.ParameterSetName) {
    'GetConstant' {
        if ($WordConstants.ContainsKey($ConstantName)) {
            return $WordConstants[$ConstantName]
        } else {
            throw "Constant '$ConstantName' not found. Use -ListAvailable to see available constants."
        }
    }
    'GetAll' {
        return $WordConstants
    }
    'ListConstants' {
        return $WordConstants.Keys | Sort-Object
    }
    default {
        # Default behavior - return all constants
        return $WordConstants
    }
}