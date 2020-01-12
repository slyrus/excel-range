
(asdf:defsystem #:excel-range
  :author "Cyrus Harmon"
  :licence "BSD"
  :depends-on ("select" "cl-ppcre")
  :serial t
  :components
  ((:file "package")
   (:file "excel-range")))

