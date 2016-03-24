(ns markdown2docx.docx-writer-test
  (:require [markdown2docx.docx-writer :refer :all]
            [clojure.test :refer :all]
            [clojure.java.io :as io]))

(deftest smoke-test
  (testing "write creates a valid file"
    (let [css {:.title {:text-align "center"}}
          document {:document
                    [{:heading [{:level 1} {:text "H1 {.title}"}]}
                     {:paragraph
                      [{:text "Emphasis, aka italics, with "}
                       {:italic [{:text "asterisks"}]}
                       {:text " or "}
                       {:italic [{:text "underscores"}]}
                       {:text "."}]}
                     {:paragraph
                      [{:text "Combined emphasis with "}
                       {:bold
                        [{:text "asterisks and "} {:italic [{:text "underscores"}]}]}
                       {:text "."}]}
                     {:ordered-list
                      [{:list-item [{:paragraph [{:text "First ordered list item"}]}]}
                       {:list-item [{:paragraph [{:text "Another item"}]}]}
                       {:list-item
                        [{:paragraph
                          [{:text
                            "Actual numbers don't matter, just that it's a number"}]}
                         {:ordered-list
                          [{:list-item [{:paragraph [{:text "Ordered sub-list"}]}]}]}]}
                       {:list-item [{:paragraph [{:text "And another item."}]}]}]}
                     {:table-block
                      [{:table-head
                        [{:table-row
                          [{:table-cell [{:text "Markdown"}]}
                           {:table-cell [{:text "Less"}]}
                           {:table-cell [{:text "Pretty"}]}]}]}
                       {:table-body
                        [{:table-row
                          [{:table-cell [{:italic [{:text "Still"}]}]}
                           {:table-cell [{:text "renders"}]}
                           {:table-cell [{:bold [{:text "nicely"}]}]}]}
                         {:table-row
                          [{:table-cell [{:text "1"}]}
                           {:table-cell [{:text "2"}]}
                           {:table-cell [{:text "3"}]}]}]}]}]}]
      (is (= true (.exists (write "sample.docx" document css)))))))
