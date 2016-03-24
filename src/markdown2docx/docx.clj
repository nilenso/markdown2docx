(ns markdown2docx.docx
  (:require [clojure.java.io :refer [file]])
  (:import [org.docx4j.openpackaging.packages WordprocessingMLPackage]
           [org.docx4j XmlUtils]
           [org.docx4j.jaxb Context]
           [org.docx4j.openpackaging.parts.WordprocessingML NumberingDefinitionsPart FooterPart]
           [org.docx4j.wml Numbering P TcPr TblWidth TblPr CTBorder STBorder TblBorders
            JcEnumeration]
           [org.docx4j.wml.PPrBase.NumPr]
           [java.math BigInteger]))


(defonce factory (Context/getWmlObjectFactory))
(defonce initialNumbering (slurp "initialNumbering.xml"))
(defonce heading {1 "Heading1"
                  2 "Heading2"
                  3 "Heading3"
                  4 "Heading4"
                  5 "Heading5"
                  6 "Heading6"})

(defonce emphasis-type {:bold #(.setB %1 %2)
                        :italic #(.setI %1 %2)})

(defonce alignment {:center JcEnumeration/CENTER
                    :left JcEnumeration/LEFT
                    :right JcEnumeration/RIGHT})

(defn create-package []
  (WordprocessingMLPackage/createPackage))

(defn maindoc [package]
  (.getMainDocumentPart package))

(defn save [package filename]
  (let [docx (file filename)]
    (.save package docx)
    docx))

(defn add-to
  [object content]
  (-> object
      (.getContent)
      (.add content)))

(defn restart-numbering
  [ndp]
  (fn [_]
    (.restart ndp 1 0 1)))

(defn set-border
  [border]
  (doto border
    (.setColor "auto")
    (.setSz (new BigInteger "4"))
    (.setSpace (new BigInteger "0"))
    (.setVal STBorder/SINGLE)))

(defn set-table-border
  [table-border border]
  (doto table-border
    (.setBottom border)
    (.setTop border)
    (.setRight border)
    (.setLeft border)
    (.setInsideH border)
    (.setInsideV border)))

(defn add-border
  [table]
  (let [border (new CTBorder)
        table-border (new TblBorders)]
    (set-border border)
    (set-table-border table-border border)
    (.setTblPr table (new TblPr))
    (-> table
      .getTblPr
      (.setTblBorders table-border))))

(defn set-cell-width
  [tablecell width]
  (let [table-cell-properties (new TcPr)
        table-width (new TblWidth)]
    (.setW table-width (BigInteger/valueOf width))
    (.setTcW table-cell-properties table-width)
    (.setTcPr tablecell table-cell-properties)))

(defn add-table-cell
  [doc]
  (let [row (:row doc)
        cell (.createTc factory)
        p (.createP factory)
        r (.createR factory)
        rpr (.createRPr factory)
        ppr (.createPPr factory)]
    (.setPPr p ppr)
    (.setRPr r rpr)
    (add-to p r)
    (add-to cell p)
    (set-cell-width cell 5500)
    (add-to row cell)
    (assoc doc :p p :r r :rpr rpr :ppr ppr)))

(defn add-table-row
  [doc]
  (let [table (:table doc)
        row (.createTr factory)]
    (add-to table row)
    (assoc doc :row row)))

(defn add-table
  [doc]
  (let [maindoc (:maindoc doc)
        table (.createTbl factory)
        tblpr (.createTblPr factory)
        jc (.createJc factory)]
    (.setVal jc JcEnumeration/CENTER)
    (.setJc tblpr jc)
    (.setTblPr table tblpr)
    (.addObject maindoc table)
    (assoc doc :table table)))

(defn add-simple-paragraph
  [maindoc text]
  (.createParagraphOfText maindoc text))

(defn add-style-heading
  [doc level]
  (let [ppr (:ppr doc)
        ppr-style (.createPPrBasePStyle factory)]
    (.setPStyle ppr ppr-style)
    (.setVal ppr-style (heading level))))

(defn add-ordered-list
  [doc]
  (let [ndp (new NumberingDefinitionsPart)
        maindoc (:maindoc doc)]
    (.addTargetPart maindoc ndp)
    (.setJaxbElement ndp (XmlUtils/unmarshalString initialNumbering))
    (assoc doc :ndp ndp)))

(defn add-text-list
  [doc numid ilvl]
  (let [ppr (:ppr doc)
        numpr (.createPPrBaseNumPr factory)
        ilvlelement (.createPPrBaseNumPrIlvl factory)
        numidelement (.createPPrBaseNumPrNumId factory)]
    (.setNumPr ppr numpr)
    (.setIlvl numpr ilvlelement)
    (.setVal ilvlelement (BigInteger/valueOf ilvl))
    (.setNumId numpr numidelement)
    (.setVal numidelement (BigInteger/valueOf numid))))

(defn set-emphasis-text
  [doc emphasis]
  (let [r (.createR factory)
        p (:p doc)
        rpr (.createRPr factory)
        b (new org.docx4j.wml.BooleanDefaultTrue)]
    (.setVal b true)
    ((emphasis emphasis-type) rpr b)
    (.setRPr r rpr)
    (add-to p r)
    (assoc doc :r r)))

(defn add-bold-italic-text
  [doc]
  (let [r (.createR factory)
        p (:p doc)
        rpr (.createRPr factory)
        b (new org.docx4j.wml.BooleanDefaultTrue)]
    (.setVal b true)
    (.setB rpr b)
    (.setI rpr b)
    (add-to p r)
    (assoc doc :r r)))

(defn add-text-align
  [doc align]
  (let [ppr (:ppr doc)
        jc (.createJc factory)]
    (.setVal jc (align alignment))
    (.setJc ppr jc)
    doc))

(defn add-text
  [doc text]
  (let [t (.createText factory)
        r (or (:r doc) (.createR factory))]
    (when (nil? (:r doc))
      (add-to (:p doc) r))
    (add-to r t)
    (.setSpace t "preserve")
    (.setValue t text)))

(defn add-paragraph
  [doc]
  (let [maindoc (:maindoc doc)
        p (.createP factory)
        ppr (.createPPr factory)]
    (.setPPr p ppr)
    (-> maindoc
        (.getJaxbElement)
        (.getBody)
        (.getContent)
        (.add p))
    (assoc doc :p p :ppr ppr)))

(defn add-pagenumber-field
  [p]
  (let [r (.createR factory)
        txt (.createText factory)]
    (doto txt
      (.setSpace "preserve")
      (.setValue " PAGE   \\* MERGEFORMAT "))
    (add-to r (.createRInstrText factory txt))
    (add-to p r)))

(defn add-field-begin
  [p]
  (let [r (.createR factory)
        fldchar (.createFldChar factory)]
    (.setFldCharType fldchar org.docx4j.wml.STFldCharType/BEGIN)
    (add-to r fldchar)
    (add-to p r)))

(defn add-field-end
  [p]
  (let [r (.createR factory)
        fldchar (.createFldChar factory)]
    (.setFldCharType fldchar org.docx4j.wml.STFldCharType/END)
    (add-to r fldchar)
    (add-to p r)))

(defn create-footer-with-page-num
  []
  (let [ftr (.createFtr factory)
        p (.createP factory)]
    (add-field-begin p)
    (add-pagenumber-field p)
    (add-field-end p)
    (add-to ftr p)
    ftr))

(defn get-section-pr
  [sections]
  (-> sections
      (.get (dec (.size sections)))
      (.getSectPr)))

(defn set-section-pr
  [sections sectpr]
  (-> sections
      (.get (dec (.size sections)))
      (.setSectPr sectpr)))

(defn create-footer-part
  [package]
  (let [footer-part (new FooterPart)]
    (.setPackage footer-part package)
    (.setJaxbElement footer-part (create-footer-with-page-num))
    (.addTargetPart (maindoc package) footer-part)))

(defn create-footer-reference
  [package relationship]
  (let [footer-reference (.createFooterReference factory)
        sections (-> package
                     (.getDocumentModel)
                     (.getSections))
        sectpr (get-section-pr sections)
        newsectpr (if (nil? sectpr)
                    (do
                      (.createSectPr factory)
                      (.addObject (maindoc package) sectpr)
                      (set-section-pr sections sectpr))
                    sectpr)]
    (doto footer-reference
      (.setId (.getId relationship))
      (.setType org.docx4j.wml.HdrFtrRef/DEFAULT))
    (-> newsectpr
        (.getEGHdrFtrReferences)
        (.add footer-reference))))
