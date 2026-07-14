"""Barreira aritmetica de capacidade antes de qualquer alteracao na planta."""

from dataclasses import dataclass

from ScannerPremissas import normalize_val


@dataclass(frozen=True)
class CapacityResult:
    vazios: int
    reducoes: int
    demanda_operacional: int
    demanda_salas: int
    reducoes_invalidas: tuple[str, ...]

    @property
    def capacidade(self):
        return self.vazios + self.reducoes

    @property
    def demanda(self):
        return self.demanda_operacional + self.demanda_salas

    @property
    def deficit(self):
        return max(0, self.demanda - self.capacidade)

    @property
    def viavel(self):
        return self.deficit == 0 and not self.reducoes_invalidas

    def aviso(self):
        linhas = [
            "PROCESSAMENTO NAO EXECUTADO: CAPACIDADE INSUFICIENTE", "",
            f"PAs 'vazio' disponiveis: {self.vazios}",
            f"Reducoes autorizadas: {self.reducoes}",
            f"Capacidade total: {self.capacidade}",
            f"Demanda operacional: {self.demanda_operacional}",
            f"Lugares em salas internas: {self.demanda_salas}",
            f"Demanda total: {self.demanda}",
            f"Deficit: {self.deficit}",
        ]
        if self.reducoes_invalidas:
            linhas.extend(["", "Reducoes invalidas:", *self.reducoes_invalidas])
        linhas.extend(["", "Nenhuma alteracao foi aplicada a planta.", "Celulas em branco nao sao consideradas PAs disponiveis."])
        return "\n".join(linhas)


def avaliar_capacidade(ws, allowed_cells, reducoes, novos_clientes, criar_ambientes, client_cells):
    vazios = sum(1 for r, c in allowed_cells if normalize_val(ws.cell(r, c).value) == "VAZIO")
    total_reducoes = sum(max(0, int(qtd or 0)) for qtd in reducoes.values())
    demanda_operacional = sum(max(0, int(nc.get("PAs", 0) or 0)) for nc in novos_clientes)
    demanda_salas = sum(max(0, int(amb.get("sala_lugares", 0) or 0)) for amb in criar_ambientes)
    invalidas = []
    for cliente, reducao in reducoes.items():
        disponivel = len(client_cells.get(cliente, set()))
        if reducao < 0 or reducao > disponivel:
            invalidas.append(f"Cliente '{cliente}': reducao {reducao}, inventario {disponivel}.")
    return CapacityResult(vazios, total_reducoes, demanda_operacional, demanda_salas, tuple(invalidas))