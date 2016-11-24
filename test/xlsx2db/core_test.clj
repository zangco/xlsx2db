(ns xlsx2db.core-test
  (:require [clojure.test :refer :all]
            [xlsx2db.core :refer :all]
            [clojure.java.jdbc :refer [quoted]]))

(deftest a-test
  (testing "Command to link table"
    (is (= "SET TABLE abc SOURCE \"abc.csv\";"
           (#'text-table-dll "abc" identity)))
    (is (= "SET TABLE \"abc\" SOURCE \"abc.csv\";"
           (#'text-table-dll "abc" (quoted \")))))
  (testing "Command to create table dll"
    (is (= "a"
           (#'table-ddl "abc" ["one" "two"] identity)
           ))))
