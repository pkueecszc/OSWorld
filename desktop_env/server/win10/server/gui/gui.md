# 接口
以下接口都有/gui前缀，具体接口中不再重复描述

POST("/click", handler.ExecGuiClick)
POST("/drag", handler.ExecGuiDrag)
POST("/left_double", handler.ExecGuiLeftDouble)
POST("/right_single", handler.ExecGuiRightSingle)
POST("/scroll", handler.ExecGuiScroll)
POST("/move_to", handler.ExecGuiMoveTo)
POST("/mouse_down", handler.ExecGuiMouseDown)
POST("/mouse_up", handler.ExecGuiMouseUp)
POST("/type", handler.ExecGuiType)
POST("/hotkey", handler.ExecGuiHotkey)
POST("/press", handler.ExecGuiPress)
POST("/release", handler.ExecGuiRelease)
POST("/wait", handler.ExecGuiWait)


# 参数
以下是go语言形式的入参定义：

type GuiClickInput struct {
	CoordinateX int `json:"coordinate_x"`
	CoordinateY int `json:"coordinate_y"`
}

type GuiDragInput struct {
	SourceCoordinateX int `json:"source_coordinate_x"`
	SourceCoordinateY int `json:"source_coordinate_y"`
	TargetCoordinateX int `json:"target_coordinate_x"`
	TargetCoordinateY int `json:"target_coordinate_y"`
}

type GuiLeftDoubleInput struct {
	CoordinateX int `json:"coordinate_x"`
	CoordinateY int `json:"coordinate_y"`
}

type GuiRightSingleInput struct {
	CoordinateX int `json:"coordinate_x"`
	CoordinateY int `json:"coordinate_y"`
}

type GuiScrollInput struct {
	CoordinateX *int   `json:"coordinate_x"`
	CoordinateY *int   `json:"coordinate_y"`
	Direction   string `json:"direction"` // "up", "down", "left", "right"
}

type GuiMoveToInput struct {
	CoordinateX int `json:"coordinate_x"`
	CoordinateY int `json:"coordinate_y"`
}

type GuiMouseDownInput struct {
	CoordinateX int    `json:"coordinate_x,omitempty"`
	CoordinateY int    `json:"coordinate_y,omitempty"`
	Button      string `json:"button"` // "left", "right", "middle"
}

type GuiMouseUpInput struct {
	CoordinateX int    `json:"coordinate_x,omitempty"`
	CoordinateY int    `json:"coordinate_y,omitempty"`
	Button      string `json:"button"` // "left", "right", "middle"
}

type GuiTypeInput struct {
	Text string `json:"text"`
}

type GuiHotkeyInput struct {
	Keys []string `json:"keys"` // e.g. ["ctrl", "c"]
}

type GuiPressInput struct {
	Key string `json:"key"`
}

type GuiReleaseInput struct {
	Key string `json:"key"`
}

type GuiWaitInput struct {
	Duration int `json:"duration"` // seconds
}


# 返回值

以下是所有gui接口共用的返回定义

type GuiOperatorResp struct {
	CommonResp
	Screenshot     string          `json:"screenshot"`
}

type CommonResp struct {
	Code    int    `json:"code"`
	Success bool   `json:"success"`
	Message string `json:"message"`
}
