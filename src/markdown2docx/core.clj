(ns markdown2docx.core
  (:require [markdown2docx.md :as md]
            [markdown2docx.docx-writer :as docx-writer]))

(defn -main [& [md-file docx-file]]
  (let [md-string (slurp md-file)
        md-map (md/parse md-string)]
    (docx-writer/write docx-file md-map)))
