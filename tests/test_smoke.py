import json
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_json(*parts):
    with open(os.path.join(ROOT, *parts), encoding="utf-8") as f:
        return json.load(f)


class DataSmokeTest(unittest.TestCase):
    def test_recipes(self):
        data = load_json("data", "raw_recipes.json")
        recipes = data.get("recipes", data) if isinstance(data, dict) else data
        self.assertIsInstance(recipes, list)
        self.assertGreaterEqual(len(recipes), 500)
        first = recipes[0]
        for key in ("id", "name", "ingredients", "steps"):
            self.assertIn(key, first)

    def test_graph(self):
        graph = load_json("data", "graph.json")
        nodes = graph.get("nodes") or graph.get("entities")
        edges = graph.get("edges") or graph.get("relations")
        self.assertIsInstance(nodes, list)
        self.assertIsInstance(edges, list)
        self.assertGreaterEqual(len(nodes), 500)
        self.assertGreaterEqual(len(edges), 5000)

    def test_experiments(self):
        summary = load_json("experiments", "summary.json")
        self.assertAlmostEqual(summary["graph"]["accuracy"], 0.885, places=3)
        self.assertAlmostEqual(summary["pure"]["accuracy"], 0.385, places=3)
        self.assertAlmostEqual(summary["graph"]["hallucination"], 0.0, places=3)
        self.assertTrue(os.path.exists(os.path.join(ROOT, "experiments", "ablation.json")))

    def test_aliases_and_taboo(self):
        aliases = load_json("data", "alias_map.json")
        pairs = load_json("data", "incompatible_pairs.json")
        self.assertIsInstance(aliases, dict)
        self.assertIsInstance(pairs, (list, dict))


if __name__ == "__main__":
    unittest.main(verbosity=2)
