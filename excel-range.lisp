
(cl:in-package #:excel-range)

(defun excel-col-index (col)
  (flet ((char-num (char)
           (1+ (- (char-code (char-upcase char)) (char-code #\A)))))
    (1- (reduce (lambda (x y)
                  (+ (* x 26) (char-num y))) col :initial-value 0))))

(defun excel-index (col row)
  (cons (1- (parse-integer row)) (excel-col-index col)))

(defun excel-cell-index (cell)
  (cl-ppcre:register-groups-bind (col row)
      ("([a-zA-Z]+)(\\d+)" cell)
    (excel-index col row)))

(defun excel-cell (data cell)
  (destructuring-bind (r . c)
      (excel-cell-index cell )
    (select:select data r c)))

(defun excel-range-indices (range)
  (cl-ppcre:register-groups-bind (start end)
      ("([a-zA-Z]+\\d+):([a-zA-Z]+\\d+)" range)
    (list (excel-cell-index start)
          (excel-cell-index end))))

(defun excel-range (data range)
  (destructuring-bind ((startr . startc) (endr . endc))
      (excel-range-indices range)
    (select:select data
      (if (eql startr endr)
          startr
          (select:range startr (1+ endr)))
      (if (eql startc endc)
          startc
          (select:range startc (1+ endc))))))

(defun excel-data (data range-or-cell)
  (if (position #\: range-or-cell)
      (excel-range data range-or-cell)
      (excel-cell data range-or-cell)))

