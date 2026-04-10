from __future__ import annotations

from werkzeug.urls import url_encode

from odoo import _, http
from odoo.exceptions import AccessError, MissingError, UserError, ValidationError
from odoo.http import content_disposition, request

from odoo.addons.portal.controllers.portal import pager as portal_pager
from odoo.addons.sale.controllers.portal import CustomerPortal


class AluminiumJoineryCustomerPortal(CustomerPortal):
    def _empty_portal_line(self, sequence=10):
        return {
            "sequence": sequence,
            "gamme_id": False,
            "serie_id": False,
            "modele_id": False,
            "qty": 1,
            "width_mm": "",
            "height_mm": "",
        }

    def _get_joinery_partner(self):
        return request.env.user.partner_id

    def _get_joinery_configuration_domain(self):
        return [("partner_id", "=", self._get_joinery_partner().id)]

    def _get_joinery_quote_domain(self):
        partner = self._get_joinery_partner()
        return [
            ("joinery_configuration_id", "!=", False),
            ("partner_id", "=", partner.id),
            ("joinery_configuration_id.partner_id", "=", partner.id),
        ]

    def _get_portal_quote(self, order_id):
        order = request.env["sale.order"].sudo().search(
            self._get_joinery_quote_domain() + [("id", "=", order_id)],
            limit=1,
        )
        if not order:
            raise MissingError(_("Ce devis n'existe pas ou n'est pas accessible."))
        return order

    def _get_portal_configuration(self, configuration_id):
        configuration = request.env["aluminium.joinery.configuration"].sudo().search(
            self._get_joinery_configuration_domain() + [("id", "=", configuration_id)],
            limit=1,
        )
        if not configuration:
            raise MissingError(_("Cette configuration n'existe pas ou n'est pas accessible."))
        return configuration

    def _get_exportable_configuration(self, configuration_id):
        user = request.env.user
        if user.has_group("base.group_user"):
            configuration = request.env["aluminium.joinery.configuration"].browse(configuration_id).exists()
            if not configuration:
                raise MissingError(_("Cette configuration n'existe pas ou n'est pas accessible."))
            configuration.check_access("read")
            return configuration
        return self._get_portal_configuration(configuration_id)

    def _prepare_home_portal_values(self, counters):
        values = super()._prepare_home_portal_values(counters)
        Configuration = request.env["aluminium.joinery.configuration"].sudo()
        SaleOrder = request.env["sale.order"].sudo()
        if "joinery_configuration_count" in counters:
            values["joinery_configuration_count"] = Configuration.search_count(self._get_joinery_configuration_domain())
        if "joinery_quote_count" in counters:
            values["joinery_quote_count"] = SaleOrder.search_count(self._get_joinery_quote_domain())
        return values

    def _prepare_configurator_catalog(self):
        Gamme = request.env["aluminium.joinery.gamme"].sudo()
        Serie = request.env["aluminium.joinery.serie"].sudo()
        Modele = request.env["aluminium.joinery.modele"].sudo()
        return {
            "gammes": Gamme.search([("active", "=", True)], order="name"),
            "series": Serie.search([("active", "=", True)], order="gamme_id, name"),
            "modeles": Modele.search([("active", "=", True)], order="serie_id, name"),
        }

    def _prepare_configurator_form_values(self, configuration=None, post=None):
        if post:
            line_indices = self._extract_line_indices(post)
            lines = []
            for index in line_indices:
                lines.append(
                    {
                        "sequence": int(post.get(f"line_sequence_{index}") or ((index + 1) * 10)),
                        "gamme_id": int(post.get(f"line_gamme_id_{index}")) if post.get(f"line_gamme_id_{index}") else False,
                        "serie_id": int(post.get(f"line_serie_id_{index}")) if post.get(f"line_serie_id_{index}") else False,
                        "modele_id": int(post.get(f"line_modele_id_{index}")) if post.get(f"line_modele_id_{index}") else False,
                        "qty": post.get(f"line_qty_{index}", "1"),
                        "width_mm": post.get(f"line_width_mm_{index}", ""),
                        "height_mm": post.get(f"line_height_mm_{index}", ""),
                    }
                )
            if not lines:
                lines = [self._empty_portal_line()]
            return {
                "configuration_id": int(post.get("configuration_id")) if post.get("configuration_id") else False,
                "project_name": post.get("project_name", ""),
                "lines": lines,
            }
        if configuration:
            return configuration._get_portal_form_values()
        return {
            "configuration_id": False,
            "project_name": "",
            "lines": [self._empty_portal_line()],
        }

    def _prepare_joinery_configurator_values(self, configuration=None, form_values=None, error=None, success=None, page=1, mode="edit"):
        values = self._prepare_portal_layout_values()
        Configuration = request.env["aluminium.joinery.configuration"].sudo()
        domain = self._get_joinery_configuration_domain()
        pager = portal_pager(
            url="/my/configurateur",
            total=Configuration.search_count(domain),
            page=page,
            step=10,
        )
        configurations = Configuration.search(domain, order="write_date desc, id desc", limit=10, offset=pager["offset"])
        page_name = "joinery_results" if mode == "results" else "joinery_configurator"
        additional_title = _("Resultats de configuration") if mode == "results" else _("Calculateur de menuiserie")
        values.update(
            {
                "page_name": page_name,
                "additional_title": additional_title,
                "configurator_catalog": self._prepare_configurator_catalog(),
                "form_values": form_values or self._prepare_configurator_form_values(configuration=configuration),
                "configuration": configuration,
                "configurations": configurations,
                "pager": pager,
                "default_url": "/my/configurateur",
                "configurator_mode": mode,
                "history_active_id": configuration.id if configuration else False,
                "summary_groups": configuration._get_report_groups("summary") if configuration and configuration.summary_ids else [],
                "result_groups": configuration._get_report_groups("result") if configuration and configuration.result_line_ids else [],
                "grouped_result_sections": configuration._get_grouped_result_sections() if configuration and configuration.summary_ids else [],
                "production_groups": configuration._get_report_groups("production") if configuration and configuration.result_line_ids else [],
                "can_generate_quote": bool(configuration and configuration.summary_ids),
                "quote_url": configuration.sale_order_id.get_joinery_portal_url() if configuration and configuration.sale_order_id else False,
                "results_url": configuration.get_portal_results_url() if configuration else False,
                "detailed_results_xlsx_url": (
                    configuration.get_detailed_results_xlsx_url() if configuration and configuration.result_line_ids else False
                ),
                "material_summary_pdf_url": (
                    configuration.get_portal_material_summary_pdf_url() if configuration and configuration.summary_ids else False
                ),
                "error_message": error,
                "success_message": success,
            }
        )
        return values

    def _extract_portal_payload(self, post):
        line_indices = self._extract_line_indices(post)
        if not line_indices:
            raise ValidationError(_("Ajoutez au moins une ligne de configuration."))
        payloads = []
        missing_messages = []
        for index in line_indices:
            fields_to_check = (
                (f"line_gamme_id_{index}", _("la gamme")),
                (f"line_serie_id_{index}", _("la serie")),
                (f"line_modele_id_{index}", _("le modele")),
                (f"line_qty_{index}", _("la quantite")),
                (f"line_width_mm_{index}", _("la largeur")),
                (f"line_height_mm_{index}", _("la hauteur")),
            )
            missing = [label for field_name, label in fields_to_check if not post.get(field_name)]
            if missing:
                missing_messages.append(_("Ligne %(line)s : renseignez %(fields)s.", line=index + 1, fields=", ".join(missing)))
                continue
            try:
                payloads.append(
                    {
                        "sequence": int(post.get(f"line_sequence_{index}") or ((index + 1) * 10)),
                        "gamme_id": int(post[f"line_gamme_id_{index}"]),
                        "serie_id": int(post[f"line_serie_id_{index}"]),
                        "modele_id": int(post[f"line_modele_id_{index}"]),
                        "qty": int(post[f"line_qty_{index}"]),
                        "width_mm": float(post[f"line_width_mm_{index}"]),
                        "height_mm": float(post[f"line_height_mm_{index}"]),
                    }
                )
            except (TypeError, ValueError) as exc:
                raise ValidationError(_("Les dimensions et quantites des lignes doivent etre numeriques.")) from exc
        if missing_messages:
            raise ValidationError(" ".join(missing_messages))
        if not payloads:
            raise ValidationError(_("Ajoutez au moins une ligne valide."))
        return {
            "project_name": post.get("project_name", ""),
            "lines": payloads,
        }

    def _has_posted_lines(self, post):
        return any(key.startswith("line_") for key in post)

    def _extract_line_indices(self, post):
        indices = set()
        for key in post:
            if not key.startswith("line_"):
                continue
            suffix = key.rsplit("_", 1)[-1]
            if suffix.isdigit():
                indices.add(int(suffix))
        return sorted(indices)

    @http.route(
        ["/my/configurateur", "/my/configurateur/page/<int:page>", "/my/configurateur/<int:configuration_id>"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_my_joinery_configurator(self, configuration_id=None, page=1, **kwargs):
        configuration = self._get_portal_configuration(configuration_id) if configuration_id else None
        success = kwargs.get("success")
        if success == "calculated":
            success = _("Le calcul a ete realise avec succes.")
        elif success == "quoted":
            success = _("Le devis a ete genere avec succes.")
        elif success == "saved":
            success = _("Le brouillon a ete enregistre avec succes.")
        elif success == "duplicated":
            success = _("La configuration a ete dupliquee dans un nouveau brouillon.")
        values = self._prepare_joinery_configurator_values(configuration=configuration, success=success, page=page)
        if configuration:
            request.session["my_joinery_config_history"] = [configuration.id] + [
                rec.id for rec in values["configurations"] if rec.id != configuration.id
            ][:99]
        return request.render("aluminium_joinery_configurator.portal_my_joinery_configurator", values)

    @http.route(
        ["/my/configurateur/<int:configuration_id>/resultats"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_my_joinery_results(self, configuration_id, **kwargs):
        configuration = self._get_portal_configuration(configuration_id)
        success = kwargs.get("success")
        error = kwargs.get("error")
        if success == "calculated":
            success = _("Le calcul a ete realise avec succes.")
        elif success == "quoted":
            success = _("Le devis a ete genere avec succes.")
        if error == "material_unavailable":
            error = _("La synthese matiere n'est pas encore disponible pour cette configuration.")
        values = self._prepare_joinery_configurator_values(
            configuration=configuration,
            success=success,
            error=error,
            mode="results",
        )
        request.session["my_joinery_config_history"] = [configuration.id] + request.session.get("my_joinery_config_history", [])[:99]
        return request.render("aluminium_joinery_configurator.portal_my_joinery_results", values)

    @http.route(
        ["/my/configurateur/<int:configuration_id>/synthese-matiere/pdf"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_joinery_material_summary_pdf(self, configuration_id, **kwargs):
        configuration = self._get_portal_configuration(configuration_id)
        if not configuration.summary_ids:
            return request.redirect(f"{configuration.get_portal_results_url()}?error=material_unavailable")

        report_service = request.env["ir.actions.report"].sudo().with_company(configuration.company_id)
        pdf_content = report_service._render_qweb_pdf(
            "aluminium_joinery_configurator.report_aluminium_joinery_material",
            res_ids=configuration.ids,
            data={"report_type": "pdf"},
        )[0]
        filename = "Synthese_matiere_%s.pdf" % (configuration.name or configuration.id)
        headers = [
            ("Content-Type", "application/pdf"),
            ("Content-Length", str(len(pdf_content))),
            ("Content-Disposition", content_disposition(filename)),
        ]
        return request.make_response(pdf_content, headers=headers)

    @http.route(
        ["/aluminium_joinery/export/detailed-results/<int:configuration_id>.xlsx"],
        type="http",
        auth="user",
        website=True,
    )
    def joinery_detailed_results_xlsx(self, configuration_id, **kwargs):
        configuration = self._get_exportable_configuration(configuration_id)
        if not configuration.result_line_ids:
            raise MissingError(_("Aucun resultat detaille n'est disponible pour cette configuration."))
        xlsx_content = configuration._build_detailed_result_xlsx()
        headers = [
            ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ("Content-Length", str(len(xlsx_content))),
            ("Content-Disposition", content_disposition(configuration._get_detailed_result_export_filename())),
        ]
        return request.make_response(xlsx_content, headers=headers)

    @http.route("/my/configurateur/calculate", type="http", auth="user", website=True, methods=["POST"])
    def portal_joinery_calculate(self, **post):
        configuration = None
        try:
            if post.get("configuration_id"):
                configuration = self._get_portal_configuration(int(post["configuration_id"]))
            if self._has_posted_lines(post):
                payload = self._extract_portal_payload(post)
                configuration = request.env["aluminium.joinery.configuration"].portal_upsert_configuration(
                    self._get_joinery_partner(), payload["lines"], configuration=configuration, project_name=payload.get("project_name")
                )
            elif not configuration:
                raise ValidationError(_("Ajoutez au moins une ligne de configuration."))
            configuration.sudo().action_calculate()
            return request.redirect(f"{configuration.get_portal_results_url()}?success=calculated")
        except (ValidationError, UserError, AccessError) as exc:
            values = self._prepare_joinery_configurator_values(
                configuration=configuration,
                form_values=self._prepare_configurator_form_values(configuration=configuration, post=post),
                error=str(exc),
            )
            return request.render("aluminium_joinery_configurator.portal_my_joinery_configurator", values)

    @http.route("/my/configurateur/save", type="http", auth="user", website=True, methods=["POST"])
    def portal_joinery_save(self, **post):
        configuration = None
        try:
            if post.get("configuration_id"):
                configuration = self._get_portal_configuration(int(post["configuration_id"]))
            payload = self._extract_portal_payload(post)
            configuration = request.env["aluminium.joinery.configuration"].portal_upsert_configuration(
                self._get_joinery_partner(), payload["lines"], configuration=configuration, project_name=payload.get("project_name")
            )
            return request.redirect(f"{configuration.get_portal_url()}?success=saved")
        except (ValidationError, UserError, AccessError) as exc:
            values = self._prepare_joinery_configurator_values(
                configuration=configuration,
                form_values=self._prepare_configurator_form_values(configuration=configuration, post=post),
                error=str(exc),
            )
            return request.render("aluminium_joinery_configurator.portal_my_joinery_configurator", values)

    @http.route("/my/configurateur/quote", type="http", auth="user", website=True, methods=["POST"])
    def portal_joinery_quote(self, **post):
        configuration = None
        try:
            if post.get("configuration_id"):
                configuration = self._get_portal_configuration(int(post["configuration_id"]))
            if self._has_posted_lines(post):
                payload = self._extract_portal_payload(post)
                configuration = request.env["aluminium.joinery.configuration"].portal_upsert_configuration(
                    self._get_joinery_partner(), payload["lines"], configuration=configuration, project_name=payload.get("project_name")
                )
                configuration.sudo().action_calculate()
            elif not configuration:
                raise ValidationError(_("Ajoutez au moins une ligne de configuration."))
            elif not configuration.summary_ids:
                configuration.sudo().action_calculate()
            configuration.sudo().action_generate_quotation()
            return request.redirect(f"{configuration.sale_order_id.get_joinery_portal_url()}?success=quoted")
        except (ValidationError, UserError, AccessError) as exc:
            values = self._prepare_joinery_configurator_values(
                configuration=configuration,
                form_values=self._prepare_configurator_form_values(configuration=configuration, post=post),
                error=str(exc),
            )
            return request.render("aluminium_joinery_configurator.portal_my_joinery_configurator", values)

    @http.route(
        "/my/configurateur/<int:configuration_id>/duplicate",
        type="http",
        auth="user",
        website=True,
        methods=["POST"],
    )
    def portal_joinery_duplicate(self, configuration_id, **post):
        configuration = self._get_portal_configuration(configuration_id)
        duplicated = configuration.sudo().portal_duplicate_for_partner(self._get_joinery_partner())
        return request.redirect(f"{duplicated.get_portal_url()}?success=duplicated")

    @http.route(
        ["/my/joinery/quotes", "/my/joinery/quotes/page/<int:page>"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_my_joinery_quotes(self, page=1, **kwargs):
        SaleOrder = request.env["sale.order"].sudo()
        domain = self._get_joinery_quote_domain()
        pager = portal_pager(
            url="/my/joinery/quotes",
            total=SaleOrder.search_count(domain),
            page=page,
            step=self._items_per_page,
        )
        quotes = SaleOrder.search(domain, order="date_order desc, id desc", limit=self._items_per_page, offset=pager["offset"])
        values = self._prepare_portal_layout_values()
        values.update(
            {
                "page_name": "joinery_quote_list",
                "additional_title": _("Mes devis"),
                "quotes": quotes,
                "pager": pager,
                "default_url": "/my/joinery/quotes",
            }
        )
        request.session["my_joinery_quotes_history"] = quotes.ids[:100]
        return request.render("aluminium_joinery_configurator.portal_my_joinery_quotes", values)

    @http.route(
        ["/my/joinery/quotes/<int:order_id>"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_my_joinery_quote_page(self, order_id, success=None, **kwargs):
        order = self._get_portal_quote(order_id)
        configuration = order.joinery_configuration_id
        values = self._prepare_portal_layout_values()
        values.update(
            {
                "page_name": "joinery_quote_detail",
                "additional_title": _("Devis menuiserie"),
                "quote": order,
                "configuration": configuration,
                "summary_groups": configuration._get_report_groups("summary") if configuration else [],
                "production_groups": configuration._get_report_groups("production") if configuration and configuration.result_line_ids else [],
                "success_message": _("Le devis a ete genere avec succes.") if success == "quoted" else False,
            }
        )
        request.session["my_joinery_quotes_history"] = [order.id] + request.session.get("my_joinery_quotes_history", [])[:99]
        return request.render("aluminium_joinery_configurator.portal_my_joinery_quote_page", values)

    @http.route(
        ["/my/configurations", "/my/configurations/page/<int:page>"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_my_joinery_configurations(self, page=1, **kwargs):
        return request.redirect("/my/configurateur%s" % (f"/page/{page}" if page and int(page) > 1 else ""))

    @http.route(
        ["/demander-un-devis", "/request-quote", "/quote-request"],
        type="http",
        auth="user",
        website=True,
    )
    def portal_joinery_quote_redirect(self, **kwargs):
        query = url_encode(kwargs) if kwargs else ""
        return request.redirect("/my/configurateur%s" % (f"?{query}" if query else ""))
