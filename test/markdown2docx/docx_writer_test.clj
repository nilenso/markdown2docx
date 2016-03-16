(ns markdown2docx.docx-writer-test
  (:require [markdown2docx.docx-writer :refer :all]
            [clojure.test :refer :all]
            [clojure.java.io :as io]))

;; TODO: write at least one very high-level test executing each type of branch

(deftest smoke-test
  (is (= "NOPE" (write "/does/not/exist" {:does-not "exist"}))))
