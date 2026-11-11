(ns build
  "Build tasks. Run with `clojure -T:build <task>`."
  (:require [clojure.tools.build.api :as b]))

(def lib 'horizontal-maker/horizontal-maker)
(def version "2.0.0")
(def class-dir "target/classes")
(def uber-file "target/horizontal-maker.jar")
(def basis (delay (b/create-basis {:project "deps.edn"})))

(defn clean
  "Remove the target directory."
  [_]
  (b/delete {:path "target"}))

(defn javac
  "Compile the Java sources under src/java into target/classes."
  [_]
  (b/javac {:src-dirs ["src/java"]
            :class-dir class-dir
            :basis @basis
            :javac-opts ["--release" "11"]}))

(defn run-tests
  "Compile Java sources and run the test suite."
  [_]
  (javac nil)
  (let [cmd (b/java-command {:basis (b/create-basis {:project "deps.edn" :aliases [:test]})
                             :main 'clojure.main
                             :main-args ["-m" "cognitect.test-runner" "-d" "tests"]})
        {:keys [exit]} (b/process cmd)]
    (when-not (zero? exit)
      (throw (ex-info "Tests failed" {:exit exit})))))

(defn uber
  "Build a standalone jar at target/horizontal-maker.jar."
  [_]
  (clean nil)
  (javac nil)
  (b/copy-dir {:src-dirs ["src/clj"] :target-dir class-dir})
  (b/compile-clj {:basis @basis
                  :ns-compile '[horizontal-maker.start]
                  :class-dir class-dir})
  (b/uber {:class-dir class-dir
           :uber-file uber-file
           :basis @basis
           :main 'horizontal-maker.start}))
