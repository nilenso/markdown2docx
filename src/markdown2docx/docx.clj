(ns markdown2docx.docx
  (:require [clojure.java.io :refer [file]])
  (:import [org.docx4j.openpackaging.packages WordprocessingMLPackage]
           [org.docx4j XmlUtils]
           [org.docx4j.jaxb Context]
           [org.docx4j.openpackaging.parts.WordprocessingML NumberingDefinitionsPart]
           [org.docx4j.wml Numbering P TcPr TblWidth TblPr CTBorder STBorder TblBorders]
           [org.docx4j.wml.PPrBase.NumPr]
           [java.math BigInteger]))


(defonce factory (Context/getWmlObjectFactory))
(defonce initialNumbering (slurp "initialNumbering.xml"))

(defn create-package []
  (WordprocessingMLPackage/createPackage))

(defn maindoc [package]
  (.getMainDocumentPart package))

(defn save [package filename]
  (.save package (file filename)))

(defn add-to
  [object content]
  (-> object
      (.getContent)
      (.add content)))

(defn restart-numbering
  [ndp]
  (fn [_]
    (.restart ndp 1 0 1)))

(defn add-text-list
  [p numid ilvl text]
  (let [t (.createText factory)
        run (.createR factory)
        ppr (.createPPr factory)
        numpr (.createPPrBaseNumPr factory)
        ilvlelement (.createPPrBaseNumPrIlvl factory)
        numidelement (.createPPrBaseNumPrNumId factory)]
    (.setValue t text)
    (add-to run t)
    (add-to p run)
    (.setPPr p ppr)
    (.setNumPr ppr numpr)
    (.setIlvl numpr ilvlelement)
    (.setVal ilvlelement (BigInteger/valueOf ilvl))
    (.setNumId numpr numidelement)
    (.setVal numidelement (BigInteger/valueOf numid))))

(defn add-ordered-list
  [maindoc]
  (let [ndp (new NumberingDefinitionsPart)]
    (.addTargetPart maindoc ndp)
    (.setJaxbElement ndp (XmlUtils/unmarshalString initialNumbering))
    ndp))

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
  [row]
  (let [cell (.createTc factory)
        p (.createP factory)]
    (add-to cell p)
    (set-cell-width cell 2500)
    (add-to row cell)
    p))

(defn add-table-row
  [table]
  (let [row (.createTr factory)]
    (add-to table row)
    row))

(defn add-table
  [maindoc]
  (let [table (.createTbl factory)]
    (add-border table)
    (.addObject maindoc table)
    table))

(defn add-simple-paragraph
  [maindoc text]
  (.createParagraphOfText maindoc text))


(defn add-text
  [p text]
  (let [r (.createR factory)
        t (.createText factory)
        ppr (.createPPr factory)
        ind (.createPPrBaseInd factory)]
    (.setInd ppr ind)
    (.setPPr p ppr)
    (add-to r t)
    (add-to p r)
    (.setValue t text)))

(defn add-paragraph
  [maindoc]
  (let [p (.createP factory)]
    (-> maindoc
        (.getJaxbElement)
        (.getBody)
        (.getContent)
        (.add p))
    p))
