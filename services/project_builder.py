from __future__ import annotations


class ProjectBuilder:
    def __init__(self, env):
        self.env = env

    def build_for_configuration(self, configuration):
        if configuration.project_project_id:
            return configuration.project_project_id
        project = self.env["project.project"].create(
            {
                "name": configuration.name,
                "partner_id": configuration.partner_id.id,
                "joinery_configuration_id": configuration.id,
            }
        )
        configuration.project_project_id = project
        return project

