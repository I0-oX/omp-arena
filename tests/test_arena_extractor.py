"""Tests del extractor vendored (port de upstream arena-mcp).

Se ejecutan en Linux sin Arena ni pywin32: solo cubren la logica pura
(operandos, expresiones, compatibilidad, cobertura). Las tools COM se
verifican en Windows con Arena instalado.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "plugins" / "arena-mcp" / "server"))

import unittest

import arena_extractor as extractor


class FakeOperand:
    def __init__(self, name, *, array=False, parent=None, default=""):
        self.Name = name
        self.Prompt = name
        self.DefaultValue = default
        self.Required = False
        self.ControlType = 0
        self.Array = array
        self.Entry = False
        self.Exit = False
        self.ParentOperand = parent or ""


class FakeDefinition:
    def __init__(self, operands):
        self.Operands = operands


class FakeModule:
    def __init__(self):
        self.Definition = FakeDefinition(
            [
                FakeOperand("Name"),
                FakeOperand("Assignments", array=True),
                FakeOperand("Type", parent="Assignments"),
                FakeOperand("Value", parent="Assignments"),
            ]
        )
        self.values = {
            "Name": "Initialize",
            "Type(1)": "Variable",
            "Value(1)": "1",
            "Type(2)": "Attribute",
            "Value(2)": "2",
        }

    def Data(self, key):
        if key not in self.values:
            raise RuntimeError("missing tuple")
        return self.values[key]


class ExtractorTests(unittest.TestCase):
    def test_repeat_group_extraction(self):
        result = extractor._extract_operands(FakeModule(), max_repeat_rows=10)

        self.assertEqual(result["scalars"][0]["value"], "Initialize")
        group = result["repeat_groups"][0]
        self.assertEqual(group["name"], "Assignments")
        self.assertEqual(len(group["rows"]), 2)
        self.assertEqual(group["rows"][1]["values"]["Type"], "Attribute")
        self.assertFalse(group["truncated"])

    def test_compatibility_classification(self):
        result = extractor._compatibility_from_modules(
            [
                {"definition": "Assign"},
                {"definition": "Assign"},
                {"definition": "Route"},
                {"definition": "VBA"},
            ]
        )

        self.assertEqual(
            result["automatic"], [{"definition": "Assign", "count": 2}]
        )
        self.assertEqual(
            result["assisted"], [{"definition": "Route", "count": 1}]
        )
        self.assertEqual(
            result["manual_or_unmapped"], [{"definition": "VBA", "count": 1}]
        )

    def test_clamp_rejects_out_of_range_value(self):
        with self.assertRaises(ValueError):
            extractor._clamp(0, 1, 10, "limit")

    def test_expression_audit_finds_scalar_and_repeat_values(self):
        module = {
            "id": "module-0001",
            "definition": "Expression",
            "caption": "Rates",
            "operands": extractor._extract_operands(
                FakeModule(), max_repeat_rows=10
            ),
        }

        result = extractor._audit_expressions([module], max_items=10)

        self.assertEqual(result["candidate_count"], 5)
        self.assertFalse(result["truncated"])
        self.assertEqual(result["candidates"][0]["module_id"], "module-0001")
        self.assertTrue(result["candidates"][0]["analysis"]["lexically_valid"])

    def test_expression_analysis_extracts_symbols_and_functions(self):
        result = extractor._analyze_expression(
            "MAX(Resource.Capacity, Variable_A[2]) + 1.5E-3"
        )

        self.assertTrue(result["lexically_valid"])
        self.assertEqual(result["functions"], ["MAX"])
        self.assertEqual(
            result["identifiers"], ["Resource.Capacity", "Variable_A"]
        )

    def test_expression_analysis_reports_unbalanced_delimiter(self):
        result = extractor._analyze_expression("NORM(10, 2")

        self.assertFalse(result["lexically_valid"])
        self.assertIn("Unclosed '(' at offset 4.", result["errors"])

    def test_external_dependency_audit_detects_file_operand(self):
        module = {
            "id": "module-0001",
            "definition": "ReadWrite",
            "caption": "Load input",
            "operands": {
                "scalars": [
                    {
                        "name": "FileName",
                        "prompt": "File Name",
                        "value": r"C:\data\input.xlsx",
                    }
                ],
                "repeat_groups": [],
            },
        }
        model = type(
            "Model",
            (),
            {
                "ExternalRef": "",
                "VisualizationFileName": "",
                "AutoPublishSimVariables": False,
            },
        )()

        result = extractor._audit_external_dependencies(
            model, [module], {"projects": []}, max_items=10
        )

        self.assertEqual(result["candidate_count"], 2)
        self.assertEqual(
            {item["kind"] for item in result["candidates"]},
            {"integration_module", "operand_reference"},
        )

    def test_coverage_report_blocks_uncaptured_vba(self):
        modules = [
            {
                "id": "module-0001",
                "definition": "VBA",
                "caption": "Hook",
                "operands": {"scalars": [], "repeat_groups": []},
            }
        ]

        result = extractor._build_coverage_report(
            modules=modules,
            module_total=1,
            collections={"Connections": {"count": 0}},
            expressions={"candidate_count": 0, "truncated": False},
            vba={
                "accessible": True,
                "project_count": 1,
                "line_count": 12,
                "captured_lines": 0,
                "source_complete": False,
            },
            submodels={"accessible": True, "count": 0},
            templates={"accessible": True, "installed_panels": []},
            dependencies={"candidate_count": 1, "truncated": False},
        )

        vba_surface = next(
            item for item in result["surfaces"] if item["key"] == "vba"
        )
        resource_surface = next(
            item
            for item in result["surfaces"]
            if item["key"] == "resources_schedules_sets"
        )
        self.assertEqual(vba_surface["status"], "metadata_only")
        self.assertEqual(resource_surface["status"], "not_present")
        self.assertEqual(result["translation_readiness"], "review_required")
        self.assertIn("VBA", result["manual_or_unmapped_definitions"])

    def test_scope_prefix_updates_module_and_connection_ids(self):
        modules = [{"id": "module-0001"}, {"id": "module-0002"}]
        connections = [
            {
                "id": "connection-0001",
                "source": {
                    "module_id": "module-0001",
                    "candidate_module_ids": ["module-0001"],
                },
                "destination": {
                    "module_id": "module-0002",
                    "candidate_module_ids": ["module-0002"],
                },
            }
        ]

        extractor._prefix_scope_ids(modules, connections, "root/submodel-0001")

        self.assertEqual(
            modules[0]["id"], "root/submodel-0001/module-0001"
        )
        self.assertEqual(
            connections[0]["destination"]["module_id"],
            "root/submodel-0001/module-0002",
        )

    def test_complete_expression_and_siman_template_coverage(self):
        result = extractor._build_coverage_report(
            modules=[],
            module_total=0,
            collections={"Connections": {"count": 0}},
            expressions={
                "candidate_count": 2,
                "truncated": False,
                "lexical_error_count_in_returned": 0,
            },
            vba={"accessible": True, "line_count": 0},
            submodels={"accessible": True, "count": 0},
            templates={
                "accessible": True,
                "installed_panels": [],
                "unresolved_definitions": [],
            },
            dependencies={"candidate_count": 0, "truncated": False},
            siman_source={"requested": True, "source_complete": True},
        )
        statuses = {item["key"]: item["status"] for item in result["surfaces"]}

        self.assertEqual(statuses["expressions"], "extracted")
        self.assertEqual(statuses["templates"], "extracted")


if __name__ == "__main__":
    unittest.main()
