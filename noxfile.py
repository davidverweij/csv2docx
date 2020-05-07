import nox

locations = "csv2docx", "tests", "noxfile.py"

@nox.session(python=["3.8"])
def tests(session):
    """Run the test suite using poetry"""
    session.run("poetry", "install", external=True)
    session.run("pytest", "--cov")

@nox.session(python=["3.8"])
def lint(session):
    args = session.posargs or locations
    session.install("flake8")
    session.run("flake8", *args)
