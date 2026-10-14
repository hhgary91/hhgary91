;;; ==========================================================================
;;; TenderSheets.lsp
;;; Builds tender detail sheets from the Standard Design Library.
;;;
;;; Command: TPSHEETS
;;;   1. Open a NEW drawing from your company title-block template (.dwt).
;;;      It must have a paper-space layout (default name "TEMPLATE") that holds
;;;      the title block, inserted at 0,0 and plotted at 1:1 in mm.
;;;   2. APPLOAD this file, type TPSHEETS, pick the CAD_Export .txt file that
;;;      was saved from Design_Library_Register.xlsx (Text, Tab delimited).
;;;   3. For every CAD Sheet No. in the export, the routine copies the template
;;;      layout, fills the title block attributes, inserts each detail DWG in
;;;      its slot at 1:n, and writes the detail title and scale under it.
;;;
;;; Detail DWGs must follow the CAD_Standard sheet: drawn 1:1 in mm, frame
;;; bottom-left at 0,0, text and dims sized for the detail scale.
;;; Test on a copy of your template first.
;;; ==========================================================================
(vl-load-com)

;;; ---------------- SETTINGS: change these to match your title block -------
(setq *tp-template-layout* "TEMPLATE")    ; layout holding the empty title block
(setq *tp-titleblock*      "TITLEBLOCK")  ; title block's block name
(setq *tp-tags*                            ; (attribute tag . value key)
  '(("PROJECT"   . project)
    ("DWG_TITLE" . title)
    ("DWG_NO"    . dwgno)
    ("SCALE"     . scale)
    ("SHEET"     . sheet)
    ("DATE"      . date)))
(setq *tp-sheet-title* "STANDARD DETAILS") ; drawing title, " (SHEET n OF N)" is added
(setq *tp-area*  '(25.0 25.0 645.0 569.0)) ; detail area on paper: x1 y1 x2 y2 (mm)
(setq *tp-cols* 3)                         ; slots across
(setq *tp-rows* 2)                         ; slots down
(setq *tp-gap*  8.0)                       ; clear margin inside each slot (mm)
(setq *tp-band* 20.0)                      ; space under each detail for its title (mm)
(setq *tp-text-h* 3.5)                     ; detail title text height (mm)
(setq *tp-layer* "TP-DETAIL-TITLE")        ; layer for detail titles
;;; -------------------------------------------------------------------------

(defun tp:split (str delim / pos out)
  (while (setq pos (vl-string-search delim str))
    (setq out (cons (substr str 1 pos) out)
          str (substr str (+ pos 1 (strlen delim)))))
  (reverse (cons str out)))

(defun tp:clean (s)
  ;; Excel wraps some fields in quotes when saving as text
  (if s (vl-string-trim " \"" s) ""))

(defun tp:nth (i lst) (tp:clean (nth i lst)))

(defun tp:pad (n) (if (< n 10) (strcat "0" (itoa n)) (itoa n)))

(defun tp:read (path / f line lines)
  (if (setq f (open path "r"))
    (progn
      (while (setq line (read-line f)) (setq lines (cons line lines)))
      (close f)
      (reverse lines))))

(defun tp:fill-titleblock (blk vals / tag key)
  (vlax-for o blk
    (if (and (= (vla-get-ObjectName o) "AcDbBlockReference")
             (= (strcase (vla-get-EffectiveName o)) (strcase *tp-titleblock*))
             (= (vla-get-HasAttributes o) :vlax-true))
      (foreach att (vlax-invoke o 'GetAttributes)
        (setq tag (strcase (vla-get-TagString att)))
        (if (setq key (cdr (assoc tag *tp-tags*)))
          (if (cdr (assoc key vals))
            (vla-put-TextString att (cdr (assoc key vals)))))))))

(defun tp:layout-exists (name / found)
  (foreach l (layoutlist)
    (if (= (strcase l) (strcase name)) (setq found T)))
  found)

(defun c:TPSHEETS (/ *error* doc path lines f project prefix recs sheets
                     n k s name lay blk details i col row slotw sloth sx sy
                     pt sc ref mn mx mt rec placed skipped warn oldecho today)
  (defun *error* (msg)
    (if oldecho (setvar "CMDECHO" oldecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort")))
      (princ (strcat "\nTPSHEETS error: " msg)))
    (princ))
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))
        oldecho (getvar "CMDECHO")
        placed 0 skipped 0)
  (setvar "CMDECHO" 0)

  (if (not (tp:layout-exists *tp-template-layout*))
    (progn
      (alert (strcat "Layout \"" *tp-template-layout* "\" not found.\n\n"
                     "Open a new drawing from your title-block template, or change\n"
                     "*tp-template-layout* at the top of TenderSheets.lsp."))
      (exit)))

  (setq path (getfiled "Select CAD_Export text file" "" "txt" 0))
  (if (not path) (exit))
  (setq lines (tp:read path))
  (if (not lines) (progn (alert "Could not read the file.") (exit)))

  ;; --- parse: header rows then one row per detail
  (setq project "" prefix "STD")
  (foreach line lines
    (setq f (tp:split line "\t"))
    (cond
      ((= (strcase (tp:nth 0 f)) "PROJECT")   (setq project (tp:nth 1 f)))
      ((= (strcase (tp:nth 0 f)) "DWGPREFIX") (setq prefix (tp:nth 1 f)))
      ((and (> (atoi (tp:nth 0 f)) 0) (/= (tp:nth 2 f) ""))
       (if (= (strcase (tp:nth 7 f)) "APPROVED")
         ;; (sheet cadfile scale title id rev)
         (setq recs (cons (list (atoi (tp:nth 0 f)) (tp:nth 2 f) (atof (tp:nth 3 f))
                                (tp:nth 4 f) (tp:nth 5 f) (tp:nth 6 f))
                          recs))
         (progn
           (setq skipped (1+ skipped))
           (princ (strcat "\n  Skipped (status " (tp:nth 7 f) "): " (tp:nth 5 f))))))))
  (setq recs (reverse recs))
  (if (not recs) (progn (alert "No Approved details found in the file.") (exit)))

  ;; --- unique sheet numbers, ascending
  (foreach r recs (if (not (member (car r) sheets)) (setq sheets (cons (car r) sheets))))
  (setq sheets (vl-sort sheets '<)
        n (length sheets) k 0
        today (menucmd "M=$(edtime,$(getvar,DATE),DD-MON-YY)"))
  (if (not (tblsearch "LAYER" *tp-layer*))
    (vla-Add (vla-get-Layers doc) *tp-layer*))
  (setq slotw (/ (- (nth 2 *tp-area*) (nth 0 *tp-area*)) *tp-cols*)
        sloth (/ (- (nth 3 *tp-area*) (nth 1 *tp-area*)) *tp-rows*))

  (foreach s sheets
    (setq k (1+ k)
          name (strcat prefix "-" (tp:pad s)))
    (if (tp:layout-exists name)
      (princ (strcat "\n  Layout " name " already exists, sheet " (itoa s) " skipped."))
      (progn
        (command "_.-LAYOUT" "_Copy" *tp-template-layout* name)
        (setq lay (vla-Item (vla-get-Layouts doc) name)
              blk (vla-get-Block lay))
        (tp:fill-titleblock blk
          (list (cons 'project project)
                (cons 'title (strcat *tp-sheet-title* " (SHEET " (itoa k) " OF " (itoa n) ")"))
                (cons 'dwgno name)
                (cons 'scale "AS SHOWN")
                (cons 'sheet (strcat (itoa k) " OF " (itoa n)))
                (cons 'date today)))
        (setq details (vl-remove-if-not '(lambda (r) (= (car r) s)) recs)
              i 0)
        (foreach rec details
          (setq col (rem i *tp-cols*)
                row (/ i *tp-cols*))
          (cond
            ((>= row *tp-rows*)
             (setq skipped (1+ skipped))
             (princ (strcat "\n  Sheet " name ": no free slot for " (nth 4 rec)
                            ". Put it on another CAD Sheet No.")))
            ((not (findfile (nth 1 rec)))
             (setq skipped (1+ skipped))
             (princ (strcat "\n  File not found: " (nth 1 rec))))
            ((<= (nth 2 rec) 0)
             (setq skipped (1+ skipped))
             (princ (strcat "\n  No detail scale in register for " (nth 4 rec))))
            (T
             (setq sx (+ (nth 0 *tp-area*) (* col slotw))
                   sy (- (nth 3 *tp-area*) (* (1+ row) sloth))
                   pt (list (+ sx *tp-gap*) (+ sy *tp-gap* *tp-band*) 0.0)
                   sc (/ 1.0 (nth 2 rec))
                   ref (vl-catch-all-apply 'vla-InsertBlock
                         (list blk (vlax-3d-point pt) (nth 1 rec) sc sc sc 0.0)))
             (if (vl-catch-all-error-p ref)
               (progn
                 (setq skipped (1+ skipped))
                 (princ (strcat "\n  Could not insert " (nth 1 rec) ": "
                                (vl-catch-all-error-message ref))))
               (progn
                 (setq placed (1+ placed))
                 ;; warn if the detail runs outside its slot
                 (vla-GetBoundingBox ref 'mn 'mx)
                 (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx))
                 (if (or (< (car mn) (+ sx 0.5)) (< (cadr mn) (+ sy *tp-band*))
                         (> (car mx) (- (+ sx slotw) 0.5)) (> (cadr mx) (- (+ sy sloth) 0.5)))
                   (princ (strcat "\n  CHECK: " (nth 4 rec) " is larger than its slot on " name
                                  ". Check the detail frame size or scale.")))
                 ;; detail title under the detail
                 (setq mt (vla-AddMText blk
                            (vlax-3d-point (list (+ sx *tp-gap*) (+ sy *tp-band* -2.0) 0.0))
                            (- slotw (* 2 *tp-gap*))
                            (strcat "{\\LDETAIL " (itoa (1+ i)) "}\\P" (nth 3 rec)
                                    "\\PSCALE 1:" (rtos (nth 2 rec) 2 0)
                                    "    " (nth 4 rec))))
                 (vla-put-Height mt *tp-text-h*)
                 (vla-put-AttachmentPoint mt acAttachmentPointTopLeft)
                 (vla-put-InsertionPoint mt (vlax-3d-point (list (+ sx *tp-gap*) (+ sy *tp-band* -2.0) 0.0)))
                 (vla-put-Layer mt *tp-layer*)))))
          (setq i (1+ i)))
        (princ (strcat "\n  Sheet " name ": " (itoa (min i (* *tp-cols* *tp-rows*))) " detail(s).")))))

  (setvar "CMDECHO" oldecho)
  (if (> n 0) (setvar "CTAB" (strcat prefix "-" (tp:pad (car sheets)))))
  (princ (strcat "\nTPSHEETS done: " (itoa placed) " detail(s) placed on " (itoa n)
                 " sheet(s), " (itoa skipped) " skipped. Read the messages above."))
  (princ))

(princ "\nTenderSheets loaded. Type TPSHEETS to build tender detail sheets.")
(princ)
