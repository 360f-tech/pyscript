import json
import os
import time

from xlcalculator import ModelCompiler, Evaluator, Model

from Singleton import SingletonService


class PremiumCalculator(metaclass=SingletonService):
    def __init__(self):
        self.model = None

    def calculate(self, proposal):
        start = time.time()
        PWD = os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
        productId = proposal["baseCoverage"]["productId"]
        formulas = f'./{productId}.json'
        model = self.createModel(formulas)
        end = time.time()
        loadModelTime = str(end - start)
        start = time.time()
        evaluator = Evaluator(model)
        self.bindingProposal(evaluator, proposal["baseCoverage"])

        sa = evaluator.evaluate("'OUTPUT'!C3")
        premium = evaluator.evaluate("'OUTPUT'!D3")
        premiumTerm = evaluator.evaluate("'OUTPUT'!F3")
        policyTerm = evaluator.evaluate("'OUTPUT'!G3")
        end = time.time()
        result = json.dumps({
            "timeLoadModel": loadModelTime,
            "premium": str(premium),
            "sum assured": str(sa),
            "policyTerm": str(policyTerm),
            "premiumTerm": str(premiumTerm),
            "PremiumTime": str(end - start)
        })
        return result

    def createModel2(self, formulas):
        if self.model is None:
            f = open(formulas, "r")
            input_dict = f.read()
            validInput = json.loads(input_dict)  # json.loads(json.dumps(input_dict))
            compiler = ModelCompiler()
            self.model = compiler.read_and_parse_dict(validInput)
        return self.model

    def createModel(self, formulas):

        if self.model is None:
            f = open(formulas, "r")
            input_dict = f.read()
            validInput = json.loads(input_dict)  # json.loads(json.dumps(input_dict))
            compiler = ModelCompiler()
            compiler.defined_names = {"'[GPP.xlsm]'!ACCM_PREM_LOOKUP": "='[GPP.xlsm]GPP(III) A'!A4:M4"}
            compiler.build_defined_names()
            # compiler.link_cells_to_defined_names()
            self.model = compiler.read_and_parse_dict(validInput)
        return self.model

    def bindingProposal(self, evaluator, coverage):
        age = coverage["parties"]["primaryInsured"]["insuredAge"]
        ppt = coverage["premiumOptions"]["premiumTerm"]
        sumInsured = coverage["sumInsured"]
        evaluator.set_cell_value("INPUT!C11", sumInsured)
        evaluator.set_cell_value("INPUT!C5", age)
        evaluator.set_cell_value("INPUT!C9", ppt)
        return evaluator

premiumCal = PremiumCalculator()

