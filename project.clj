(defproject dk.ative/docjure "1.14.0-SNAPSHOT"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/mjul/docjure"
  :license {:name "MIT License"
            :url "http://http://en.wikipedia.org/wiki/MIT_License"}
  :dependencies [[org.clojure/clojure "1.10.0"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]]
  :plugins [[lein-difftest "2.0.0"]]
  :profiles {:1.5   {:dependencies [[org.clojure/clojure "1.5.1"]]}
             :1.6   {:dependencies [[org.clojure/clojure "1.6.0"]]}
             :1.7   {:dependencies [[org.clojure/clojure "1.7.0"]]}
             :1.8   {:dependencies [[org.clojure/clojure "1.8.0"]]}
             :1.9   {:dependencies [[org.clojure/clojure "1.9.0"]]}
             :1.10  {:dependencies [[org.clojure/clojure "1.10.0"]]}
             :test  {:global-vars  {*warn-on-reflection* false}
                     :dependencies [[com.cemerick/pomegranate "0.3.0"]]}}
  :aliases {"all" ["with-profile" "1.5:1.6:1.7:1.8:1.9:1.10"]}
  :global-vars {*warn-on-reflection* true})

