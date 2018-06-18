;初始化中的获取插入点
(setq PlotPoint (getpoint "选择插入点"))

;画矩形
(setq offsetH 5 offsetV 10)
(setq Point2 (list (+ (car PlotPoint) offsetH) (+ (car (cdr PlotPoint)) offsetV)))
(command "rectang" PlotPoint Point2)

;画mtext矩形
(setq Point3 (list (+ (car PlotPoint) offsetH) (+ (car (cdr PlotPoint)) (* 0.5 offsetV))))
(command "mtext" PlotPoint Point3 "\\pxqc;line1\\Pline2" "")
(setq shux (entget (entlast)))
(setq shux (subst (cons 40 1) (assoc 40 shux) shux)) ;设定文字高度
(setq shux (subst (cons 44 1) (assoc 44 shux) shux)) ;设定行距比例
(setq shux (subst (cons 71 4) (assoc 71 shux) shux)) ;设定附着点为左中，垂直对齐
(entmod shux)

;移动PlotPoint，为下一个图形做准备
(setq PlotPoint (list (+ (car PlotPoint) offsetH) (car (cdr PlotPoint))))