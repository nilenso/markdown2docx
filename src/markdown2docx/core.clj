(ns markdown2docx.core
  (:require [markdown2clj.core :as md]
            [markdown2docx.docx-writer :as docx-writer]
            [markdown2docx.css :as css]))

(defn -main [& [md-file docx-file css-file]]
  (let [md-string (slurp md-file)
        md-map (md/parse md-string true)
        css-string (if (not (nil? css-file))
                     (slurp css-file)
                     "")
        css-map (css/parse-css css-string)]
    (when (.exists (docx-writer/write docx-file md-map css-map))
      (println "The docx file has been created @ " docx-file))))
