(ns markdown2docx.core
  (:require [clojure.java.io :refer [file]])
  (:import [org.docx4j.openpackaging.packages WordprocessingMLPackage]))

(def md-file "/Users/steven/work/nilenso/cooperative-agreement/template.md")
(def docx-file "/Users/steven/work/nilenso/cooperative-agreement/template.docx")

(defn foo
  "I don't do a whole lot."
  [x]
  (println x "Hello, World!"))

(defn zig []
  (let [pkg (WordprocessingMLPackage/createPackage)]
    (do (-> pkg
            (.getMainDocumentPart)
            (.addParagraphOfText "HAXED BY CHINESE!")))
    (.save pkg (file docx-file))))

(defn -main [& args]
  ;; TODO: take in params: markdown, css, docx
  ;; TODO: read markdown, look for css ids / classes
  ;; TODO: read in css, match classes
  ;; TODO: write to docx using corresponding docx styles for each css style
  (zig)
  )
